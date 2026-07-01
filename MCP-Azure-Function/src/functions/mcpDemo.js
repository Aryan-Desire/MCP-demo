const { app } = require('@azure/functions');
const { DefaultAzureCredential } = require('@azure/identity');
const { AIProjectClient } = require('@azure/ai-projects');

app.http('mcpDemo', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log(`Http function processed request for url "${request.url}"`);

        const aiProjectsEndpoint = process.env.AZURE_AI_PROJECTS_ENDPOINT || "https://ai-agent-mcp-resource.services.ai.azure.com/api/projects/ai-agent-mcp";
        const agentName = process.env.AZURE_AI_PROJECTS_AGENT_NAME || "mcp-demo-agent";
        const agentVersion = process.env.AZURE_AI_PROJECTS_AGENT_VERSION || "3";

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
            prompt = `Using the SharePoint MCP tool, retrieve the item with ID ${data.itemId} from the SharePoint list with ID '${data.listId}'. Once retrieved, generate a concise, professional summary of the list item data (including school name, selected school, assignee, document status, damages, and comments).`;
        } else if (!prompt && data) {
            prompt = `Please generate a concise, professional summary of the data.`
        }

        if (!prompt) {
            prompt = "List the first 5 items from the SharePoint site or tell me what lists exist in the site.";
        }

        let summary = '';
        let statusMessage = '';

        if (aiProjectsEndpoint && agentName && agentVersion) {
            try {
                context.log('Initializing AIProjectClient...');
                const projectClient = new AIProjectClient(aiProjectsEndpoint, new DefaultAzureCredential());
                const openAIClient = projectClient.getOpenAIClient();

                context.log('Creating conversation with initial user message...');
                const conversation = await openAIClient.conversations.create({
                    items: [{ type: "message", role: "user", content: prompt }]
                });
                context.log(`Created conversation (id: ${conversation.id})`);

                context.log('Generating response using agent reference...');
                const response = await openAIClient.responses.create(
                    {
                        conversation: conversation.id,
                    },
                    {
                        body: { agent: { name: agentName, version: agentVersion, type: "agent_reference" } },
                    },
                );

                summary = response.output_text || 'No output text returned by Agent.';
                statusMessage = 'Successfully executed Agent call via AIProjectClient.';
                context.log('Agent response generated successfully.');
            } catch (err) {
                context.log('Error calling AI Project Agent:', err);
                statusMessage = `Error calling Project Agent: ${err.message || String(err)}`;
                summary = `Failed to get response due to error: ${err.message || String(err)}`;
            }
        } else {
            context.log('Required configuration (aiProjectsEndpoint, agentName, or agentVersion) is missing.');
            statusMessage = 'Configuration missing. Check your environment variables.';
            summary = 'Fallback: AI Project Agent configuration is missing.';
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
