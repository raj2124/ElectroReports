// The Graph client the app uses: the real one, or (GRAPH_USE_MOCK) the real client running on the in-memory Graph.
const { createGraphClient } = require('./graphClient');
const { createMockGraph, seedDemo } = require('./graphMock');
const defaultConfig = require('./config');

function createGraph({ config = defaultConfig.graph } = {}) {
  if (!config.useMock) return createGraphClient({ config });
  const mock = createMockGraph();
  seedDemo(mock, config);
  const client = createGraphClient({ http: mock.http, config: { ...config, tenantId: 'mock', clientId: 'mock', clientSecret: 'mock', certPath: '' } });
  return { ...client, mock };
}

module.exports = { createGraph };
