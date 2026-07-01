const { app } = require('@azure/functions');

app.http('mcpDemo', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log(`Http function processed request for url "${request.url}"`);

        const openAiEndpoint = process.env.AZURE_OPENAI_ENDPOINT;
        const openAiKey = process.env.AZURE_OPENAI_API_KEY;
        const openAiDeployment = process.env.AZURE_OPENAI_DEPLOYMENT_NAME;

        if (request.method === 'POST') {
            try {
                const data = await request.json();
                
                context.log('====================================');
                context.log('RECEIVED FORM DATA FOR SUMMARY:');
                context.log(`School Name: ${data.schoolName || 'N/A'}`);
                context.log(`Selected School: ${data.selectedSchool || 'N/A'}`);
                context.log(`Assign To (User Login): ${data.assignToLogin || 'N/A'}`);
                context.log(`Assign To (Email): ${data.assignToEmail || 'N/A'}`);
                context.log(`Document Status: ${data.documentStatus || 'N/A'}`);
                context.log(`Non-Pecuniary Damages: ${data.nonPecuniaryDamages || 'N/A'}`);
                context.log(`Punitive Damages: ${data.punitiveDamages || 'N/A'}`);
                context.log(`Comments: ${data.comments || 'N/A'}`);
                context.log('====================================');

                let summary = '';
                let statusMessage = '';

                if (openAiEndpoint && openAiKey && openAiDeployment && 
                    !openAiEndpoint.startsWith('YOUR_') && 
                    !openAiKey.startsWith('YOUR_') && 
                    !openAiDeployment.startsWith('YOUR_')) {
                    try {
                        const prompt = `Please generate a concise, professional summary of the following form data:
- School Name: ${data.schoolName || 'N/A'}
- Selected School: ${data.selectedSchool || 'N/A'}
- Assigned To: ${data.assignToLogin || 'N/A'} (${data.assignToEmail || 'N/A'})
- Document Status: ${data.documentStatus || 'N/A'}
- Non-Pecuniary Damages: ${data.nonPecuniaryDamages || 'N/A'}
- Punitive Damages: ${data.punitiveDamages || 'N/A'}
- Comments: ${data.comments || 'N/A'}`;

                        const url = `${openAiEndpoint.replace(/\/$/, '')}/openai/deployments/${openAiDeployment}/chat/completions?api-version=2025-01-01-preview`;
                        
                        context.log(`Calling Azure OpenAI Chat Completions API at: ${url}`);
                        const response = await fetch(url, {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json',
                                'api-key': openAiKey
                            },
                            body: JSON.stringify({
                                messages: [
                                    {
                                        role: "system",
                                        content: "You are an assistant that creates professional summaries of case/form submissions."
                                    },
                                    {
                                        role: "user",
                                        content: prompt
                                    }
                                ],
                                max_tokens: 800,
                                temperature: 0.7
                            })
                        });

                        if (response.ok) {
                            const resJson = await response.json();
                            summary = resJson.choices?.[0]?.message?.content || 'No summary returned by Azure OpenAI.';
                            statusMessage = 'Successfully generated summary using Azure OpenAI.';
                            context.log('Summary successfully generated.');
                        } else {
                            const errText = await response.text();
                            throw new Error(`Azure OpenAI returned ${response.status}: ${errText}`);
                        }
                    } catch (err) {
                        context.log('Error generating summary via Azure OpenAI:', err);
                        statusMessage = `Error generating summary: ${err.message || String(err)}`;
                        summary = 'Failed to generate summary due to API or network error.';
                    }
                } else {
                    context.log('Azure OpenAI credentials or endpoint are not fully configured.');
                    statusMessage = 'Azure OpenAI not configured. Check environment variables.';
                    summary = `Fallback Summary (Azure OpenAI not configured): School: ${data.schoolName || 'N/A'} - Status: ${data.documentStatus || 'N/A'}`;
                }

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
            } catch (err) {
                context.log('Failed to parse request JSON:', err);
                return {
                    status: 400,
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        success: false,
                        error: 'Invalid JSON payload'
                    })
                };
            }
        }

        const name = request.query.get('name') || 'world';
        return { body: `Hello, ${name}! Send a POST request with form data to process it.` };
    }
});
