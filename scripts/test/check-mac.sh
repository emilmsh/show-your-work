#!/bin/sh
# Checks the Mac's look-up (scripts/show.js) against the demo files.
#
#   sh scripts/test/check-mac.sh           looks every case up; shows nothing
#   sh scripts/test/check-mac.sh show      then opens one source of each kind
#                                          at its words, four seconds apart
#
# Needs Word, Excel and PowerPoint. The first run asks, once per program,
# whether the terminal may control it: answer OK. Send the whole output back.
cd "$(dirname "$0")/../.." || exit 1
SYW_ROOT="$PWD" /usr/bin/osascript -l JavaScript scripts/test/check-mac.js "$@"
