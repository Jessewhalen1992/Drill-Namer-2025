#!/usr/bin/env bash
set -euxo pipefail

# Install latest LTS .NET SDK into $HOME/dotnet
curl -SL https://dot.net/v1/dotnet-install.sh | bash /dev/stdin \
  --channel LTS \
  --install-dir "$HOME/dotnet"

# Make dotnet/msbuild visible to later steps
echo 'export PATH=$HOME/dotnet:$PATH' >> "$BASH_ENV"
