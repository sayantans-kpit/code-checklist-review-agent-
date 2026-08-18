#!/bin/bash
# Uses Node 22 via nvm (pinned in .nvmrc) — system Node 12 is bypassed.
set -e

# Activate the Node version specified in .nvmrc
export NVM_DIR="$HOME/.nvm"
. "$NVM_DIR/nvm.sh"
nvm use        # reads .nvmrc → v22.22.2

echo "Node: $(node --version)  npm: $(npm --version)"

echo "Installing dependencies..."
npm install

echo "Installing vsce..."
npm install --save-dev @vscode/vsce

echo "Compiling TypeScript..."
npx tsc -p .

echo "Packaging extension..."
npx vsce package --no-dependencies

echo ""
echo "✅ Done! Install the .vsix with:"
echo "   VS Code → Extensions → ⋮ → Install from VSIX..."
echo "   OR:  code --install-extension code-review-checklist-agent-*.vsix"
