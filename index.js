const msal = require('@azure/msal-node');

const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: process.env.CLOUD_INSTANCE + process.env.TENANT_ID,
        clientSecret: process.env.CLIENT_SECRET
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: 'debug',
        },
    },
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

const tokenRequest = {
    scopes: ["https://graph.microsoft.com/.default"],
};

cca.acquireTokenByClientCredential(tokenRequest).then((response) => {
    console.log(JSON.stringify(response));

    // こっちはMSAL Node により自動的にメモリキャッシュされたトークンを取る
    cca.acquireTokenByClientCredential(tokenRequest).then((response) => {
        console.log(JSON.stringify(response));
        }).catch((error) => {
            console.log(JSON.stringify(error));
        });

}).catch((error) => {
    console.log(JSON.stringify(error));
});
