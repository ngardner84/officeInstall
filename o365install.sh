#!/bin/bash
#
# Microsoft 365 + Teams Installer — Jamf policy deployment
# Runs as root via jamf binary. No user interaction.
#
# Notes:
#   - Microsoft 365 suite installer is a Universal .pkg (Intel + Apple Silicon).
#   - New Teams for Mac is also a Universal .pkg (osx-x64 + osx-arm64).
#   - Arch detection is kept for logging/inventory purposes only.
#

set -o pipefail

LOG="/var/log/m365_install.log"
exec > >(/usr/bin/tee -a "$LOG") 2>&1
echo "=== $(date) — Starting M365 install ==="
echo "Script reached machine and began execution."

# Microsoft 365 suite — Universal .pkg
# Ref: https://go.microsoft.com/fwlink/?linkid=525133
OFFICE_SUITE_URL="https://go.microsoft.com/fwlink/?linkid=525133"

# Microsoft Teams — Universal .pkg (enterprise deployment URL)
TEAMS_URL="https://statics.teams.cdn.office.net/production-osx/enterprise/universal/MicrosoftTeams.pkg"

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
    local min_size_mb="${4:-50}"
    local max_attempts=3
    local attempt=1

    while [ $attempt -le $max_attempts ]; do
        echo "Downloading $label (attempt $attempt/$max_attempts)..."
        /usr/bin/curl -sSL \
            --http1.1 \
            --fail \
            --retry 5 \
            --retry-delay 10 \
            --retry-all-errors \
            --retry-max-time 900 \
            --connect-timeout 30 \
            -o "$dest" \
            "$url"
        local rc=$?

        if [ $rc -eq 0 ] && [ -s "$dest" ]; then
            local size_mb
            size_mb=$(/usr/bin/du -m "$dest" | /usr/bin/cut -f1)
            if [ "$size_mb" -ge "$min_size_mb" ]; then
                echo "Downloaded $label (${size_mb}MB)"
                return 0
            else
                echo "WARN: $label download is only ${size_mb}MB (expected >=${min_size_mb}MB). Retrying..."
            fi
        else
            echo "WARN: curl exited with code $rc on attempt $attempt. Retrying..."
        fi

        rm -f "$dest"
        attempt=$((attempt + 1))
        sleep 15
    done

    echo "ERROR: Failed to download $label after $max_attempts attempts."
    return 1
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
    echo "Detected architecture: $ARCH (informational only; both installers are universal)"

    OFFICE_PKG="$WORKDIR/Microsoft365.pkg"
    TEAMS_PKG="$WORKDIR/MicrosoftTeams.pkg"

    downloadPkg "$OFFICE_SUITE_URL" "$OFFICE_PKG" "Microsoft 365 Suite" 1000 || { cleanup; exit 1; }
    downloadPkg "$TEAMS_URL"        "$TEAMS_PKG"  "Microsoft Teams"      200  || { cleanup; exit 1; }

    verifyPkg "$OFFICE_PKG" "Microsoft 365 Suite" || { cleanup; exit 1; }
    verifyPkg "$TEAMS_PKG"  "Microsoft Teams"     || { cleanup; exit 1; }

    installPkg "$OFFICE_PKG" "Microsoft 365 Suite"
    OFFICE_RC=$?

    installPkg "$TEAMS_PKG" "Microsoft Teams"
    TEAMS_RC=$?

    cleanup

    if [ $OFFICE_RC -eq 0 ] && [ $TEAMS_RC -eq 0 ]; then
        echo "=== $(date) — M365 install completed successfully ==="
        /usr/local/bin/jamf recon
        wait
        exit 0
    else
        echo "=== $(date) — M365 install completed with errors ==="
        wait
        exit 1
    fi
}

main
