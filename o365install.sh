#!/bin/bash
#
# Microsoft 365 Installer — Jamf policy deployment
# Runs as root via jamf binary. No user interaction.
#

set -o pipefail

LOG="/var/log/m365_install.log"
exec >> "$LOG" 2>&1
echo "=== $(date) — Starting M365 install ==="

# Microsoft 365 suite — Universal .pkg (Intel + Apple Silicon)
OFFICE_SUITE_URL="https://go.microsoft.com/fwlink/?linkid=525133"

# Microsoft Teams — arch-specific
TEAMS_ARM64_URL="https://go.microsoft.com/fwlink/?linkid=2249065"
TEAMS_X64_URL="https://go.microsoft.com/fwlink/?linkid=2249062"

WORKDIR="/private/tmp/m365_install"
mkdir -p "$WORKDIR"

detectArch () {
    if [[ "$(/usr/bin/arch)" == "arm64" ]] || [[ "$(/usr/sbin/sysctl -n machdep.cpu.brand_string)" == *"Apple"* ]]; then
        echo "arm64"
    else
        echo "x86_64"
    fi
}

downloadPkg () {
    local url="$1"
    local dest="$2"
    local label="$3"
    echo "Downloading $label..."
    /usr/bin/curl -sSL --fail --retry 3 --retry-delay 5 -o "$dest" "$url"
    if [ $? -ne 0 ] || [ ! -s "$dest" ]; then
        echo "ERROR: Failed to download $label from $url"
        return 1
    fi
    echo "Downloaded $label ($(du -h "$dest" | cut -f1))"
}

verifyPkg () {
    local pkg="$1"
    local label="$2"
    if /usr/sbin/pkgutil --check-signature "$pkg" | grep -q "Developer ID Installer: Microsoft Corporation"; then
        echo "$label signature verified."
        return 0
    else
        echo "ERROR: $label signature check failed."
        return 1
    fi
}

installPkg () {
    local pkg="$1"
    local label="$2"
    echo "Installing $label..."
    /usr/sbin/installer -pkg "$pkg" -target / -verboseR
    local rc=$?
    if [ $rc -eq 0 ]; then
        echo "$label installed successfully."
    else
        echo "ERROR: $label installer exited with code $rc."
    fi
    return $rc
}

cleanup () {
    rm -rf "$WORKDIR"
}

main () {
    if [ "$EUID" -ne 0 ]; then
        echo "ERROR: Must run as root. Exiting."
        exit 1
    fi

    ARCH=$(detectArch)
    echo "Detected architecture: $ARCH"

    if [ "$ARCH" = "arm64" ]; then
        TEAMS_URL="$TEAMS_ARM64_URL"
    else
        TEAMS_URL="$TEAMS_X64_URL"
    fi

    OFFICE_PKG="$WORKDIR/Microsoft365.pkg"
    TEAMS_PKG="$WORKDIR/MicrosoftTeams.pkg"

    downloadPkg "$OFFICE_SUITE_URL" "$OFFICE_PKG" "Microsoft 365 Suite" || { cleanup; exit 1; }
    downloadPkg "$TEAMS_URL" "$TEAMS_PKG" "Microsoft Teams" || { cleanup; exit 1; }

    verifyPkg "$OFFICE_PKG" "Microsoft 365 Suite" || { cleanup; exit 1; }
    verifyPkg "$TEAMS_PKG" "Microsoft Teams" || { cleanup; exit 1; }

    installPkg "$OFFICE_PKG" "Microsoft 365 Suite"
    OFFICE_RC=$?

    installPkg "$TEAMS_PKG" "Microsoft Teams"
    TEAMS_RC=$?

    cleanup

    if [ $OFFICE_RC -eq 0 ] && [ $TEAMS_RC -eq 0 ]; then
        echo "=== $(date) — M365 install completed successfully ==="
        /usr/local/bin/jamf recon
        exit 0
    else
        echo "=== $(date) — M365 install completed with errors ==="
        exit 1
    fi
}

main
