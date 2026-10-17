#!/usr/bin/env bash
# build.sh - Multi-target cross compilation script for Linux/macOS
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
export TEMP="${TEMP:-$SCRIPT_DIR/.tmp}"
export TMP="${TMP:-$SCRIPT_DIR/.tmp}"
export NUGET_PACKAGES="${NUGET_PACKAGES:-$SCRIPT_DIR/.nuget/packages}"

RIDS=("win-x86" "win-x64" "linux-x64" "linux-arm" "linux-arm64")

for RID in "${RIDS[@]}"; do
    echo "Publishing for $RID..."
    dotnet publish "$SCRIPT_DIR/src/CrossReportRunner/CrossReportRunner.csproj" \
        -c Release \
        -r "$RID" \
        --self-contained true \
        -p:PublishSingleFile=true \
        -p:IncludeNativeLibrariesForSelfExtract=true \
        -p:EnableCompressionInSingleFile=true \
        -o "$SCRIPT_DIR/bin/$RID"
done
echo "All targets published to bin/ successfully."
