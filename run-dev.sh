#!/usr/bin/env zsh
export PATH="$(pwd)/.tools/node-v25.8.0-darwin-arm64/bin:$PATH"
npm run dev -- --host 0.0.0.0 --port 4173
