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
            prompt = `First, use Sharepoint MCP tool to retrieve the list of all files in the SharePoint document library named "MCP_Templates" from sharepoint site ${sharepointSiteUrl}. This will give you the actual file. Then, for the files found in the library, retrieve their text content using the getFileContent tool and generate a detailed summary of the content and data within those files.`;
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

                // Add custom getFileContent tool
                openAiTools.push({
                    type: "function",
                    function: {
                        name: "getFileContent",
                        description: "Retrieve the text content of a file (e.g. .docx, .pdf, .xlsx, .txt) from a SharePoint site using its server-relative URL",
                        parameters: {
                            type: "object",
                            properties: {
                                fileUrl: {
                                    type: "string",
                                    description: "The server-relative URL of the file (e.g. /sites/AIDevelopment/MCP_Templates/Template.docx)"
                                }
                            },
                            required: ["fileUrl"]
                        }
                    }
                });

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
                const maxLoops = 5;

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
                            if (name === "getListItems") {
                                context.log(`Intercepting getListItems for list: ${args.listTitle}...`);
                                try {
                                    const { ClientSecretCredential } = require('@azure/identity');
                                    const tenantId = process.env.TENANT_ID || process.env.AZURE_TENANT_ID;
                                    const clientId = process.env.CLIENT_ID || process.env.AZURE_CLIENT_ID;
                                    const clientSecret = process.env.CLIENT_SECRET || process.env.AZURE_CLIENT_SECRET;

                                    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
                                    const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");
                                    const graphToken = tokenResponse.token;

                                    const graphHeaders = {
                                        'Authorization': `Bearer ${graphToken}`,
                                        'Accept': 'application/json'
                                    };

                                    let siteId = process.env.SHAREPOINT_SITE_ID;
                                    if (args.url) {
                                        try {
                                            const urlObj = new URL(args.url);
                                            const hostname = urlObj.hostname;
                                            const pathname = urlObj.pathname.replace(/\/$/, "");
                                            context.log(`Resolving site ID dynamically for ${hostname}:${pathname}...`);
                                            const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${hostname}:${pathname}`, {
                                                headers: graphHeaders
                                            });
                                            if (siteRes.ok) {
                                                const siteData = await siteRes.json();
                                                siteId = siteData.id;
                                                context.log(`Resolved site ID dynamically: ${siteId}`);
                                            } else {
                                                context.log(`Failed to resolve site ID dynamically: ${siteRes.statusText}`);
                                            }
                                        } catch (urlErr) {
                                            context.log(`Error parsing url for dynamic site ID lookup:`, urlErr);
                                        }
                                    }

                                    context.log(`Fetching drives for site: ${siteId}`);
                                    const drivesRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
                                        headers: graphHeaders
                                    });
                                    if (!drivesRes.ok) {
                                        throw new Error(`Failed to fetch drives: ${drivesRes.statusText}`);
                                    }
                                    const drivesData = await drivesRes.json();
                                    const drives = drivesData.value || [];
                                    const drive = drives.find(d => d.name.toLowerCase() === args.listTitle.toLowerCase() || (d.webUrl && d.webUrl.toLowerCase().endsWith('/' + args.listTitle.toLowerCase())));
 
                                    if (!drive) {
                                        throw new Error(`List '${args.listTitle}' is not a document library or drive, falling back to standard MCP tool.`);
                                    }
                                    let items = [];
                                    context.log(`Fetching children for drive: ${drive.id}`);
                                    const itemsRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${drive.id}/root/children?$expand=listItem`, {
                                        headers: graphHeaders
                                    });
                                    if (itemsRes.ok) {
                                        const itemsData = await itemsRes.json();
                                        const sitePath = new URL(sharepointSiteUrl).pathname;
                                        items = (itemsData.value || []).map(i => {
                                            const fileUrl = `${sitePath}/${args.listTitle}/${i.name}`;
                                            const fields = i.listItem?.fields || {};
                                            return {
                                                id: i.id,
                                                fileName: i.name,
                                                fileUrl: fileUrl,
                                                ...fields
                                            };
                                        });
                                    }
 
                                    const fieldsList = new Set(["id", "fileName", "fileUrl"]);
                                    items.forEach(item => {
                                        Object.keys(item).forEach(k => fieldsList.add(k));
                                    });
 
                                    toolResult = JSON.stringify({
                                        content: [{
                                            type: "text",
                                            text: JSON.stringify({
                                                listTitle: args.listTitle,
                                                totalItems: items.length,
                                                fields: Array.from(fieldsList),
                                                items: items
                                            }, null, 2)
                                        }]
                                    });
                                } catch (err) {
                                    context.log(`Error in getListItems interception:`, err);
                                    // Fallback to standard MCP call
                                    const callRes = await mcpClient.callTool({
                                        name: name,
                                        arguments: args
                                    });
                                    toolResult = JSON.stringify(callRes);
                                }
                            } else if (name === "getFileContent") {
                                context.log(`Intercepting getFileContent for file: ${args.fileUrl}`);
                                try {
                                     const { ClientSecretCredential } = require('@azure/identity');
                                     const officeParser = require('officeparser');

                                     const tenantId = process.env.TENANT_ID || process.env.AZURE_TENANT_ID;
                                     const clientId = process.env.CLIENT_ID || process.env.AZURE_CLIENT_ID;
                                     const clientSecret = process.env.CLIENT_SECRET || process.env.AZURE_CLIENT_SECRET;

                                     const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
                                     context.log(`Requesting Microsoft Graph token...`);
                                     const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");
                                     const graphToken = tokenResponse.token;

                                     const graphHeaders = {
                                         'Authorization': `Bearer ${graphToken}`,
                                         'Accept': 'application/json'
                                     };

                                     let siteId = process.env.SHAREPOINT_SITE_ID;
                                     const match = args.fileUrl.match(/^\/(sites\/[^\/]+)/);
                                     if (match) {
                                         const sitePath = '/' + match[1];
                                         try {
                                             const urlObj = new URL(sharepointSiteUrl);
                                             const hostname = urlObj.hostname;
                                             context.log(`Resolving site ID dynamically for getFileContent: ${hostname}:${sitePath}...`);
                                             const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`, {
                                                 headers: graphHeaders
                                             });
                                             if (siteRes.ok) {
                                                 const siteData = await siteRes.json();
                                                 siteId = siteData.id;
                                                 context.log(`Resolved site ID dynamically for getFileContent: ${siteId}`);
                                             } else {
                                                 context.log(`Failed to resolve site ID dynamically for getFileContent: ${siteRes.statusText}`);
                                             }
                                         } catch (urlErr) {
                                             context.log(`Error parsing url for dynamic site ID lookup in getFileContent:`, urlErr);
                                         }
                                     }

                                     context.log(`Fetching drives for site: ${siteId}`);
                                     const drivesRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
                                         headers: graphHeaders
                                     });
                                     if (!drivesRes.ok) {
                                         throw new Error(`Failed to fetch drives: ${drivesRes.statusText}`);
                                     }
                                     const drivesData = await drivesRes.json();
                                     const drives = drivesData.value || [];
 
                                     let targetDrive = null;
                                     let relativePath = '';
                                     const normFileUrl = args.fileUrl.replace(/\\/g, '/');
 
                                     for (const drive of drives) {
                                         if (drive.webUrl) {
                                             const drivePath = new URL(drive.webUrl).pathname.replace(/\\/g, '/');
                                             if (normFileUrl.toLowerCase().startsWith(drivePath.toLowerCase())) {
                                                 targetDrive = drive;
                                                 relativePath = normFileUrl.substring(drivePath.length);
                                                 break;
                                             }
                                         }
                                     }
 
                                     if (!targetDrive) {
                                         throw new Error(`Could not find a SharePoint drive matching file URL: ${args.fileUrl}`);
                                     }
 
                                     if (!relativePath.startsWith('/')) {
                                         relativePath = '/' + relativePath;
                                     }
 
                                     const graphDownloadUrl = `https://graph.microsoft.com/v1.0/drives/${targetDrive.id}/root:${encodeURI(relativePath)}:/content`;
                                     context.log(`Directly downloading file content from Graph API: ${graphDownloadUrl}`);
 
                                     const fileResponse = await fetch(graphDownloadUrl, {
                                         headers: {
                                             'Authorization': `Bearer ${graphToken}`
                                         }
                                     });
 
                                     if (!fileResponse.ok) {
                                         throw new Error(`Failed to download file from Graph: ${fileResponse.statusText}`);
                                     }

                                    const arrayBuffer = await fileResponse.arrayBuffer();
                                    const fileBuffer = Buffer.from(arrayBuffer);
                                    const fileName = args.fileUrl.split('/').pop() || 'file';

                                    let fileText = '';
                                    if (fileName.toLowerCase().endsWith('.txt')) {
                                        fileText = fileBuffer.toString('utf8');
                                    } else {
                                        const ast = await officeParser.parseOffice(fileBuffer);
                                        fileText = await ast.toText();
                                    }

                                    toolResult = JSON.stringify({
                                        success: true,
                                        fileName: fileName,
                                        content: fileText
                                    });
                                } catch (err) {
                                    context.log(`Error intercepting getFileContent:`, err);
                                    toolResult = JSON.stringify({
                                        success: false,
                                        error: err.message || String(err)
                                    });
                                }
                            } else {
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
