import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createIWorkServer } from "../../src/server.js";

export interface TestContext {
  server: McpServer;
  client: Client;
  cleanup: () => Promise<void>;
}

export async function createTestServer(): Promise<TestContext> {
  const server = createIWorkServer({ platform: "darwin", name: "iwork-mcp-test", version: "0.0.0" });

  const client = new Client({ name: "test-client", version: "0.0.0" });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();

  await server.connect(serverTransport);
  await client.connect(clientTransport);

  return {
    server,
    client,
    async cleanup() {
      await client.close();
      await server.close();
    },
  };
}
