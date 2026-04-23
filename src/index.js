const { app } = require('@azure/functions');

app.http('GetLocations', {
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        return {
            jsonBody: {
                success: true,
                locations: [
                    { id: 1, name: 'Range A', status: 'Open' },
                    { id: 2, name: 'Range B', status: 'Reserved' },
                    { id: 3, name: 'Range C', status: 'Maintenance' }
                ]
            }
        };
    }
});
