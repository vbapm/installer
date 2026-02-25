#!/bin/sh

# Uninstaller for vbapm on macOS
# Reverses the steps performed by install.sh

set -e

lib_dir="$HOME/Library/Application Support/vbapm"
bin_dir="$lib_dir/bin"
addins_link="$HOME/vbapm Add-ins"
export_bin="export PATH=\"\$PATH:$bin_dir\""
zprofile="$HOME/.zprofile"
profile="$HOME/.profile"
bash_profile="$HOME/.bash_profile"

echo "[1/3] Removing vbapm from PATH..."
for rc_file in "$zprofile" "$profile" "$bash_profile"; do
  if [ -f "$rc_file" ] && grep -qF "$bin_dir" "$rc_file"; then
    # Use a temp file to avoid in-place issues on macOS (no sed -i without extension)
    tmp=$(mktemp)
    grep -vF "$export_bin" "$rc_file" > "$tmp"
    mv "$tmp" "$rc_file"
    echo "Removed PATH entry from $rc_file."
  fi
done

echo "[2/3] Removing add-ins link..."
if [ -L "$addins_link" ] || [ -e "$addins_link" ]; then
  rm -f "$addins_link"
  echo "Removed \"$addins_link\"."
else
  echo "Link not found, skipping."
fi

echo "[3/3] Removing vbapm installation directory..."
if [ -d "$lib_dir" ]; then
  rm -rf "$lib_dir"
  echo "Removed \"$lib_dir\"."
else
  echo "\"$lib_dir\" not found, skipping."
fi

echo ""
echo "\033[32mDone!\033[m vbapm has been uninstalled."
echo ""
echo "Note: If you had enabled \"Trust access to the VBA project object model\""
echo "in Excel > Preferences > Security, you may want to uncheck it manually."
