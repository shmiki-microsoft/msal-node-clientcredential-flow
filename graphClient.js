/*
https://learn.microsoft.com/ja-jp/graph/tutorials/node?tutorial-step=3
*/

const graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

const getAuthenticatedClient = (msalClient, scopes) => {
  if (!msalClient) {
    throw new Error('Invalid MSAL state. Client: missing');
  }
  if (!scopes || !Array.isArray(scopes) || scopes.length === 0) {
    throw new Error('Scopes must be a non-empty array.');
  }

  // getAccessTokenメソッドを持つオブジェクトをauthProviderに渡す
  const authProvider = {
    getAccessToken: async () => {
      const response = await msalClient.acquireTokenByClientCredential({
        scopes: scopes,
      });
      return response.accessToken;
    }
  };

  return graph.Client.initWithMiddleware({
    authProvider
  });
};

const getUsersTop10 = async (msalClient, scopes) => {
  const client = getAuthenticatedClient(msalClient, scopes);
  return client
    .api('/users')
    .select('displayName,mail,userPrincipalName')
    .top(10)
    .get();
};

const getUsersAll = async (msalClient, scopes) => {
  const client = getAuthenticatedClient(msalClient, scopes);
  const users = [];
  const response = await client
    .api('/users')
    .select('displayName,mail,userPrincipalName')
    .top(1)
    .get();

  // PageIterator を使って全ページ取得
  const pageIterator = new graph.PageIterator(
    client,
    response,
    (user) => {
      users.push(user);
      return true; // 続行
    }
  );
  await pageIterator.iterate();
  return { value: users };
};

// クライアントクレデンシャルフローでは /me エンドポイントは利用できません
const getUserDetails = async () => {
  throw new Error('getUserDetails is not supported in client credential flow.');
};

module.exports = {
  getUsersTop10,
  getUsersAll,
};