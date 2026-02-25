#!/bin/sh

# Based heavily on the approach used in https://github.com/denoland/deno_install
# Copyright 2018 the Deno authors. All rights reserved. MIT license.

# Installer for vbapm on macOS

set -e

lib_dir="$HOME/Library/Application Support/vbapm"
bin_dir="$lib_dir/bin"
exe="$bin_dir/vba"
addins_dir="$lib_dir/addins/build"
addins_link="$HOME/vbapm Add-ins"
zip_file="$lib_dir/vbapm-mac.tar.gz"
export_bin="export PATH=\"\$PATH:$bin_dir\""
zprofile="$HOME/.zprofile"
profile="$HOME/.profile"
bash_profile="$HOME/.bash_profile"

if [ -d "$lib_dir" ]; then existing_install=0; else existing_install=1; fi

if [ $# -eq 0 ]; then
	# Use the GitHub API to find the latest release asset URL.
	# The releases HTML page requires JavaScript to render asset links,
	# so we query the API which returns JSON instead.
	release_uri=$(curl -sSf https://api.github.com/repos/vbapm/core/releases/latest |
		grep -o '"browser_download_url": *"[^"]*vbapm-mac\.tar\.gz"' |
		head -n 1 |
		sed 's/.*"browser_download_url": *"\([^"]*\)"/\1/')
	if [ ! "$release_uri" ]; then echo "Error: Could not find release download URL."; exit 1; fi
else
	release_uri="https://github.com/vbapm/core/releases/download/${1}/vbapm-mac.tar.gz"
fi

if [ ! -d "$lib_dir" ]; then
  echo "Creating lib directory $lib_dir"
	mkdir -p "$lib_dir"
fi

echo "[1/4] Downloading vbapm..."
curl -fL# -o "$zip_file" "$release_uri"

echo "[2/4] Extracting vbapm..."
tar -xzf "$zip_file" --directory "$lib_dir"
# Strip Windows CRLF line endings from shell scripts in case the release was built on Windows
sed -i '' 's/\r//' "$bin_dir/vbapm"
sed -i '' 's/\r//' "$bin_dir/vba"
chmod +x "$bin_dir/vbapm"
chmod +x "$bin_dir/vba"
chmod +x "$lib_dir/vendor/node"

# Add bin to .zprofile / .profile / .bash_profile
echo "[3/4] Adding vbapm to PATH"
if [ -a $zprofile ] && ! grep -q "$bin_dir" $zprofile; then
  echo $export_bin >> "$zprofile"
fi
if ! [ -a $profile ] || ! grep -q "$bin_dir" $profile; then
  echo $export_bin >> "$profile"
fi
if [ -a $bash_profile ] && ! grep -q "$bin_dir" $bash_profile; then
  echo $export_bin >> "$bash_profile"
fi

# Create accessible add-ins folder
echo "[4/4] Creating link to add-ins at \"$addins_link\"..."
ln -sf "$addins_dir" "$addins_link"

echo ""
echo "\033[32mSuccess!\033[m vbapm was installed successfully."
echo ""
echo "Command: \"$exe\""
echo "Add-ins: \"$addins_link\""
echo ""
echo "Open a new Terminal window and run 'vba --help' to get started"

if (( existing_install == 0 )); then
	echo ""
	echo "[Additional Instructions]"
	echo ""
	echo "For more recent versions of Office for Mac, you will need to"
	echo "trust access to the VBA project object model"
	echo "for vbapm to work correctly."
	echo ""
	echo "1. Open Excel"
	echo "2. Click \"Excel\" in the menu bar"
	echo "3. Select \"Preferences\" in the menu"
	echo "4. Click \"Security\" in the Preferences dialog"
	echo "5. Check \"Trust access to the VBA project object model\""
	echo "   in the Security dialog"
fi
