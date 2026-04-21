# Base image: Node 20 on Alpine Linux (Alpine is a minimal Linux distro — keeps the image small)
FROM node:20-alpine

# Set the working directory inside the container.
# All subsequent commands run relative to this path.
WORKDIR /app

# ── Install presentation-engine dependencies ────────────────────────────────
# We copy package.json BEFORE the source code to exploit Docker's layer cache:
# if the source changes but package.json doesn't, Docker skips re-running npm install.
COPY presentation-engine/package.json ./presentation-engine/package.json
RUN cd presentation-engine && npm install --omit=dev

# Copy the three scripts (no build step needed, plain JS)
COPY presentation-engine/src ./presentation-engine/src

# ── Install MCP server dependencies ─────────────────────────────────────────
COPY services/mcp-pptx/package.json ./services/mcp-pptx/package.json
RUN cd services/mcp-pptx && npm install --omit=dev

# Copy the MCP server entry point
COPY services/mcp-pptx/index.js ./services/mcp-pptx/index.js

# ── Runtime ──────────────────────────────────────────────────────────────────
# Document that the container listens on port 8080 (IBM Code Engine will map this)
EXPOSE 8080

# Start the server. CMD is the default command when the container runs.
CMD ["node", "services/mcp-pptx/index.js"]
