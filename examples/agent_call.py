"""Launch iwork-mcp over stdio and call a pair of read-only tools."""

import asyncio
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SERVER_ENTRY = PROJECT_ROOT / "dist" / "index.js"
SAFE_TOOLS = ("numbers_list_templates", "keynote_list_themes")


async def main() -> None:
    if not SERVER_ENTRY.is_file():
        raise SystemExit("Server is not built. Run `npm install` and `npm run build` first.")

    server = StdioServerParameters(command="node", args=[str(SERVER_ENTRY)])

    async with stdio_client(server) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            response = await session.list_tools()
            available = {tool.name for tool in response.tools}
            print(f"Server exposes {len(available)} tools.")

            for tool_name in SAFE_TOOLS:
                if tool_name not in available:
                    raise RuntimeError(f"Expected tool is unavailable: {tool_name}")

                result = await session.call_tool(tool_name, arguments={})
                print(f"\n{tool_name}:")
                for item in result.content:
                    text = getattr(item, "text", None)
                    print(text if text is not None else item)


if __name__ == "__main__":
    asyncio.run(main())
