const { app } = require('@azure/functions');

app.http('mcpDemo', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log(`Http function processed request for url "${request.url}"`);

        if (request.method === 'POST') {
            try {
                const data = await request.json();
                
                context.log('====================================');
                context.log('RECEIVED FORM DATA:');
                context.log(`School Name: ${data.schoolName || 'N/A'}`);
                context.log(`Selected School: ${data.selectedSchool || 'N/A'}`);
                context.log(`Assign To (User Login): ${data.assignToLogin || 'N/A'}`);
                context.log(`Assign To (Email): ${data.assignToEmail || 'N/A'}`);
                context.log(`Document Status: ${data.documentStatus || 'N/A'}`);
                context.log(`Non-Pecuniary Damages: ${data.nonPecuniaryDamages || 'N/A'}`);
                context.log(`Punitive Damages: ${data.punitiveDamages || 'N/A'}`);
                context.log(`Comments: ${data.comments || 'N/A'}`);
                context.log(`Attachments: ${Array.isArray(data.attachments) ? data.attachments.join(', ') : 'None'}`);
                context.log('====================================');

                return {
                    status: 200,
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        success: true,
                        message: 'Form data received and logged successfully in Azure Function!',
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
