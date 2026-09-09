#!/bin/bash

# Script to connect to Cisco IOS switches via SSH, execute show commands,
# and capture output in a single file with per-switch demarcation.
# This script prompts for username and password once at execution.
# WARNING: Using sshpass for password authentication is insecure as the password
# may be visible in process listings or logs. For production use, prefer SSH key-based
# authentication to avoid password exposure. sshpass must be installed on the system.
# Usage: ./script.sh <input_file>
# Example: ./script.sh ips.txt

# NOTE: set -e is intentionally NOT used here so that SSH failures on individual
# hosts do not abort the entire script. Errors are handled per-host explicitly.
set -uo pipefail

# ── Colours ──────────────────────────────────────────────────────────────────
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
NC='\033[0m'  # No Colour

# ── Tunables ─────────────────────────────────────────────────────────────────
MAX_RETRIES=2          # How many times to retry a failed SSH before giving up
SSH_CONNECT_TIMEOUT=10 # Seconds before SSH connection attempt is aborted
SSH_CMD_TIMEOUT=60     # Hard wall-clock timeout (seconds) for the entire SSH session

# ── Logging helpers ───────────────────────────────────────────────────────────
log_info()    { echo -e "${CYAN}INFO:${NC}  $(date '+%Y-%m-%d %H:%M:%S') - $*" >&2; }
log_success() { echo -e "${GREEN}OK:${NC}    $(date '+%Y-%m-%d %H:%M:%S') - $*" >&2; }
log_warn()    { echo -e "${YELLOW}WARN:${NC}  $(date '+%Y-%m-%d %H:%M:%S') - $*" >&2; }
log_error()   { echo -e "${RED}ERROR:${NC} $(date '+%Y-%m-%d %H:%M:%S') - $*" >&2; }

# ── Dependency check ──────────────────────────────────────────────────────────
for dep in sshpass ssh ping timeout; do
    if ! command -v "$dep" &>/dev/null; then
        log_error "Required command '$dep' not found. Please install it and retry."
        exit 1
    fi
done

# ── Argument validation ───────────────────────────────────────────────────────
if [ "$#" -ne 1 ]; then
    log_error "Invalid arguments. Usage: $0 <input_file>"
    exit 1
fi

input_file="$1"

if [ ! -f "$input_file" ] || [ ! -r "$input_file" ]; then
    log_error "Input file '$input_file' does not exist or is not readable."
    exit 1
fi

# ── Timestamped output files ──────────────────────────────────────────────────
timestamp=$(date '+%Y%m%d_%H%M%S')
output_file="output_${timestamp}.txt"
failed_file="failed_hosts_${timestamp}.txt"

# ── Credentials ───────────────────────────────────────────────────────────────
read -rp "Enter SSH username: " username
log_info "Username entered: $username"

read -rsp "Enter SSH password: " password
echo ""
log_info "Password entered (length: ${#password})"

if [ -z "$password" ]; then
    log_error "Password cannot be empty."
    exit 1
fi

# ── Prepare output files ──────────────────────────────────────────────────────
> "$output_file"
> "$failed_file"
log_info "Output file: $output_file"

# ── Counters ──────────────────────────────────────────────────────────────────
total=0
succeeded=0
failed=0
skipped=0

# ── SSH command function with retry ──────────────────────────────────────────
run_ssh() {
    local ip="$1"
    local attempt=1

    while [ "$attempt" -le "$MAX_RETRIES" ]; do
        if [ "$attempt" -gt 1 ]; then
            log_warn "Retry $((attempt - 1))/$((MAX_RETRIES - 1)) for $ip ..."
            sleep 2
        fi

        timeout "$SSH_CMD_TIMEOUT" \
            sshpass -p "$password" ssh \
                -o StrictHostKeyChecking=no \
                -o ConnectTimeout="$SSH_CONNECT_TIMEOUT" \
                -o BatchMode=no \
                -o ServerAliveInterval=10 \
                -o ServerAliveCountMax=3 \
                "$username@$ip" << 'ENDSSH'
terminal length 0
show running-config | include (hostname|snmp-server location)
show mac address-table
show interface description
show arp
show cdp neighbors
show ip route connected
sh version
show interfaces stat
exit
ENDSSH
        local rc=$?
        if [ "$rc" -eq 0 ]; then
            return 0
        fi

        log_warn "SSH to $ip failed (exit code $rc), attempt $attempt/$MAX_RETRIES."
        attempt=$((attempt + 1))
    done

    return 1
}

# ── Main loop ─────────────────────────────────────────────────────────────────
while IFS= read -r ip; do
    # Strip whitespace and carriage returns (handles CRLF files)
    ip="${ip//[$'\t\r ']/}"

    # Skip empty lines
    if [ -z "$ip" ]; then
        continue
    fi

    # Skip comment lines
    if [[ "$ip" == \#* ]]; then
        log_info "Skipping comment: $ip"
        continue
    fi

    total=$((total + 1))
    log_info "[$total] Processing: $ip"

    # ── Pre-flight ping ───────────────────────────────────────────────────────
    if ! ping -c 1 -W 2 "$ip" &>/dev/null; then
        log_warn "$ip is unreachable (ping failed). Skipping."
        skipped=$((skipped + 1))
        echo "$ip  [UNREACHABLE - ping failed]" >> "$failed_file"
        {
            echo "----- Switch: $ip -----"
            echo "Timestamp: $(date '+%Y-%m-%d %H:%M:%S')"
            echo "STATUS: UNREACHABLE (ping failed)"
            echo "----- End of Switch: $ip -----"
            echo ""
        } >> "$output_file"
        continue
    fi

    # ── Header in output ──────────────────────────────────────────────────────
    {
        echo "================================================================="
        echo " Switch : $ip"
        echo " Time   : $(date '+%Y-%m-%d %H:%M:%S')"
        echo "================================================================="
        echo ""
    } >> "$output_file"

    # ── SSH with retry ────────────────────────────────────────────────────────
    log_info "Connecting to $username@$ip ..."
    if run_ssh "$ip" >> "$output_file" 2>&1; then
        log_success "$ip — commands collected successfully."
        succeeded=$((succeeded + 1))
        status="SUCCESS"
    else
        log_error "$ip — failed after $MAX_RETRIES attempt(s)."
        failed=$((failed + 1))
        status="FAILED"
        echo "$ip  [SSH FAILED after $MAX_RETRIES attempt(s)]" >> "$failed_file"
    fi

    # ── Footer in output ──────────────────────────────────────────────────────
    {
        echo ""
        echo "STATUS: $status"
        echo "================================================================="
        echo " End of Switch: $ip"
        echo "================================================================="
        echo ""
    } >> "$output_file"

done < "$input_file"

# ── Summary ───────────────────────────────────────────────────────────────────
echo ""
echo -e "${CYAN}══════════════════ SUMMARY ══════════════════${NC}" >&2
echo -e "  Total processed : $total" >&2
echo -e "  ${GREEN}Succeeded${NC}        : $succeeded" >&2
echo -e "  ${RED}Failed${NC}           : $failed" >&2
echo -e "  ${YELLOW}Skipped${NC} (unreachable): $skipped" >&2
echo -e "  Output file      : $output_file" >&2
if [ "$failed" -gt 0 ] || [ "$skipped" -gt 0 ]; then
    echo -e "  Failed hosts     : $failed_file" >&2
fi
echo -e "${CYAN}═════════════════════════════════════════════${NC}" >&2
