const { app } = require('@azure/functions');
const { AzureOpenAI } = require('openai');

app.http('mcpDemo', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log(`Http function processed request for url "${request.url}"`);

        const openAiEndpoint = process.env.AZURE_OPENAI_ENDPOINT;
        const openAiKey = process.env.AZURE_OPENAI_API_KEY;
        const openAiDeployment = process.env.AZURE_OPENAI_DEPLOYMENT_NAME;
        const openAiApiVersion = process.env.AZURE_OPENAI_API_VERSION || '2025-03-01-preview';
        const sharepointSiteUrl = process.env.SHAREPOINT_SITE_URL;

        let prompt = request.query.get('query') || request.query.get('prompt');
        let data = null;

        if (request.method === 'POST') {
            try {
                data = await request.json();
                if (data && data.query) {
                    prompt = data.query;
                } else if (data && data.prompt) {
                    prompt = data.prompt;
                }
            } catch (err) {
                context.log('Error parsing JSON body:', err);
            }     
        }

        // If no prompt was explicitly provided but we have listId and itemId, instruct AI to fetch the item via MCP and summarize
        if (!prompt && data && data.listId && data.itemId) {
            prompt = `Using the SharePoint MCP tool, retrieve all items from the SharePoint list named "MCP_demo" from sharepoint site ${sharepointSiteUrl}. Once retrieved, generate a concise, professional summary of the list item data (including school name, selected school, assignee, document status, damages, and comments).`;
        } else if (!prompt && data) {
            prompt = `Please generate a concise, professional summary of the data.`
        }

        if (!prompt) {
            prompt = "List the first 5 items from the SharePoint site or tell me what lists exist in the site.";
        }

        let summary = '';
        let statusMessage = '';

        if (openAiEndpoint && openAiKey && openAiDeployment) {
            try {
                // Dynamic import of MCP Client and Transport (CommonJS support)
                const { Client } = await import('@modelcontextprotocol/sdk/client/index.js');
                const { StdioClientTransport } = await import('@modelcontextprotocol/sdk/client/stdio.js');

                // Launch custom SharePoint MCP server locally
                const transport = new StdioClientTransport({
                    command: "node",
                    args: ["D:/Custom-MCP/server-sharepoint/build/src/index.js"],
                    env: {
                        ...process.env,
                        SHAREPOINT_CLIENT_ID: process.env.CLIENT_ID || process.env.AZURE_CLIENT_ID,
                        SHAREPOINT_CLIENT_SECRET: process.env.CLIENT_SECRET || process.env.AZURE_CLIENT_SECRET,
                        M365_TENANT_ID: process.env.TENANT_ID || process.env.AZURE_TENANT_ID,
                        AZURE_APPLICATION_ID: process.env.AZURE_CLIENT_ID,
                        AZURE_APPLICATION_CERTIFICATE_THUMBPRINT: process.env.AZURE_APPLICATION_CERTIFICATE_THUMBPRINT,
                        AZURE_APPLICATION_CERTIFICATE_PASSWORD: process.env.AZURE_APPLICATION_CERTIFICATE_PASSWORD
                    }
                });

                const mcpClient = new Client({
                    name: "mcpDemo-client",
                    version: "1.0.0"
                });

                context.log("Connecting to local Custom SharePoint MCP Server...");
                await mcpClient.connect(transport);

                context.log("Fetching tools from local MCP Server...");
                const mcpToolsResponse = await mcpClient.listTools();
                const mcpTools = mcpToolsResponse.tools || [];
                context.log(`Discovered ${mcpTools.length} tools.`);

                // Map MCP tools to standard OpenAI function tool schemas
                const openAiTools = mcpTools.map(t => ({
                    type: "function",
                    function: {
                        name: t.name,
                        description: t.description,
                        parameters: t.inputSchema
                    }
                }));

                const messages = [
                    { 
                        role: "system", 
                        content: "You are a professional assistant. You have access to tools that query a SharePoint site. Use them to fetch details and answer the query." 
                    },
                    { role: "user", content: prompt }
                ];

                const openai = new AzureOpenAI({
                    apiKey: openAiKey,
                    apiVersion: openAiApiVersion,
                    endpoint: openAiEndpoint
                });

                let summaryText = '';
                let loopCount = 0;
                const maxLoops = 2;

                while (loopCount < maxLoops) {
                    context.log(`Calling Azure OpenAI (iteration ${loopCount + 1})...`);
                    const chatCompletion = await openai.chat.completions.create({
                        model: openAiDeployment,
                        messages: messages,
                        tools: openAiTools.length > 0 ? openAiTools : undefined
                    });

                    const responseMessage = chatCompletion.choices?.[0]?.message;
                    if (!responseMessage) {
                        break;
                    }

                    messages.push(responseMessage);

                    if (responseMessage.tool_calls && responseMessage.tool_calls.length > 0) {
                        for (const toolCall of responseMessage.tool_calls) {
                            const name = toolCall.function.name;
                            const args = JSON.parse(toolCall.function.arguments);
                            context.log(`Model requested tool call: ${name} with args: ${JSON.stringify(args)}`);

                            let toolResult;
                            try {
                                const callRes = await mcpClient.callTool({
                                    name: name,
                                    arguments: args
                                });
                                toolResult = JSON.stringify(callRes);
                            } catch (toolErr) {
                                context.log(`Error executing tool ${name}:`, toolErr);
                                toolResult = `Error: ${toolErr.message || String(toolErr)}`;
                            }

                            messages.push({
                                role: "tool",
                                tool_call_id: toolCall.id,
                                content: toolResult
                            });
                        }
                    } else {
                        summaryText = responseMessage.content || '';
                        break;
                    }

                    loopCount++;
                }

                // Clean up MCP connection
                await mcpClient.close();

                summary = summaryText || 'No response generated by OpenAI.';
                statusMessage = 'Successfully executed Local Custom MCP Tool loop.';
                context.log('Local MCP tool loop completed successfully.');
            } catch (err) {
                context.log('Error executing local MCP tool loop:', err);
                statusMessage = `Error executing local MCP loop: ${err.message || String(err)}`;
                summary = `Failed to get response due to error: ${err.message || String(err)}`;
            }
        } else {
            context.log('Required Azure OpenAI configuration (endpoint, key, or deployment) is missing.');
            statusMessage = 'Configuration missing. Check your environment variables.';
            summary = 'Fallback: Azure OpenAI configuration is missing.';
        }

        if (request.method === 'POST') {
            return {
                status: 200,
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    success: true,
                    message: summary,
                    summary: summary,
                    statusMessage: statusMessage,
                    receivedData: data
                })
            };
        } else {
            return {
                status: 200,
                headers: { 'Content-Type': 'text/plain' },
                body: summary
            };
        }
    }
});
