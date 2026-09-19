#!/bin/zsh
set -euo pipefail

script_dir="${0:A:h}"
project_root="${script_dir:h}"
release_version="${1:-3.2.3}"
if [[ ! "${release_version}" =~ ^[0-9]+\.[0-9]+(\.[0-9]+)?$ ]]; then
    print -u2 "Invalid version"
    exit 1
fi
output_root="${project_root}/build/macos-${release_version}"
helper_work="${output_root}/helper-work"
helper_dist="${output_root}/helper-dist"
app_path="${output_root}/ReturnBot.app"
dmg_stage="${output_root}/dmg-stage"
dmg_path="${output_root}/ReturnBot-${release_version}-arm64.dmg"
asset_catalog="${output_root}/AppIcon.xcassets"
asset_output="${output_root}/icon-output"
icon_partial_plist="${output_root}/AppIconPartial.plist"
swift_cache="/tmp/returnbot-swift-cache"
clang_cache="/tmp/returnbot-clang-cache"

cd "${project_root}"
export PYINSTALLER_CONFIG_DIR="${output_root}/pyinstaller-cache"

if [[ ! -x "${project_root}/.venv/bin/pyinstaller" ]]; then
    print -u2 "Missing .venv/bin/pyinstaller"
    exit 1
fi

mkdir -p "${output_root}" "${swift_cache}" "${clang_cache}"
rm -rf "${helper_work}" "${helper_dist}" "${app_path}" "${dmg_stage}" "${dmg_path}" "${asset_catalog}" "${asset_output}" "${icon_partial_plist}"

icon_set="${asset_catalog}/AppIcon.appiconset"
mkdir -p "${icon_set}" "${asset_output}"
cp "${project_root}/packaging/AppIconContents.json" "${icon_set}/Contents.json"
sips -z 16 16 "${project_root}/ReturnBotIcon-flat.png" --out "${icon_set}/icon_16x16.png" >/dev/null
sips -z 32 32 "${project_root}/ReturnBotIcon-flat.png" --out "${icon_set}/icon_16x16@2x.png" >/dev/null
sips -z 32 32 "${project_root}/ReturnBotIcon-flat.png" --out "${icon_set}/icon_32x32.png" >/dev/null
sips -z 64 64 "${project_root}/ReturnBotIcon-flat.png" --out "${icon_set}/icon_32x32@2x.png" >/dev/null
sips -z 128 128 "${project_root}/ReturnBotIcon-flat.png" --out "${icon_set}/icon_128x128.png" >/dev/null
sips -z 256 256 "${project_root}/ReturnBotIcon-flat.png" --out "${icon_set}/icon_128x128@2x.png" >/dev/null
sips -z 256 256 "${project_root}/ReturnBotIcon-flat.png" --out "${icon_set}/icon_256x256.png" >/dev/null
sips -z 512 512 "${project_root}/ReturnBotIcon-flat.png" --out "${icon_set}/icon_256x256@2x.png" >/dev/null
sips -z 512 512 "${project_root}/ReturnBotIcon-flat.png" --out "${icon_set}/icon_512x512.png" >/dev/null
sips -z 1024 1024 "${project_root}/ReturnBotIcon-flat.png" --out "${icon_set}/icon_512x512@2x.png" >/dev/null
xcrun actool \
    --compile "${asset_output}" \
    --platform macosx \
    --minimum-deployment-target 26.0 \
    --app-icon AppIcon \
    --output-partial-info-plist "${icon_partial_plist}" \
    "${asset_catalog}"

"${project_root}/.venv/bin/python" -m PyInstaller \
    --noconfirm \
    --clean \
    --workpath "${helper_work}" \
    --distpath "${helper_dist}" \
    "${project_root}/ReturnBotHelper.spec"

env \
    SWIFTPM_MODULECACHE_OVERRIDE="${swift_cache}" \
    CLANG_MODULE_CACHE_PATH="${clang_cache}" \
    swift build -c release --disable-sandbox

mkdir -p "${app_path}/Contents/MacOS" "${app_path}/Contents/Resources"
cp "${project_root}/.build/release/ReturnBotMac" "${app_path}/Contents/MacOS/ReturnBot"
cp -R "${helper_dist}/ReturnBotHelper" "${app_path}/Contents/Resources/ReturnBotHelper"
cp "${asset_output}/AppIcon.icns" "${app_path}/Contents/Resources/MyIcon.icns"
cp "${project_root}/packaging/Info.plist" "${app_path}/Contents/Info.plist"
"${project_root}/.venv/bin/python" - "${app_path}/Contents/Info.plist" "${release_version}" <<'PYINFO'
import plistlib, sys
from pathlib import Path
p = Path(sys.argv[1]); info = plistlib.loads(p.read_bytes())
info['CFBundleShortVersionString'] = info['CFBundleVersion'] = sys.argv[2]
assert info.get('SUPublicEDKey'), 'Missing update signing public key'
p.write_bytes(plistlib.dumps(info))
PYINFO
mkdir -p "${app_path}/Contents/Frameworks"
ditto "${project_root}/.build/artifacts/sparkle/Sparkle/Sparkle.xcframework/macos-arm64_x86_64/Sparkle.framework" "${app_path}/Contents/Frameworks/Sparkle.framework"
chmod 755 "${app_path}/Contents/MacOS/ReturnBot" "${app_path}/Contents/Resources/ReturnBotHelper/ReturnBotHelper"

codesign --force --sign - \
    --entitlements "${project_root}/packaging/ReturnBot.entitlements" \
    "${app_path}/Contents/Resources/ReturnBotHelper/ReturnBotHelper"
codesign --force --sign - \
    --entitlements "${project_root}/packaging/ReturnBot.entitlements" \
    "${app_path}"
codesign --verify --deep --strict --verbose=2 "${app_path}"

mkdir -p "${dmg_stage}"
cp -R "${app_path}" "${dmg_stage}/ReturnBot.app"
ln -s /Applications "${dmg_stage}/Applications"
hdiutil create \
    -volname "退料機器人 ${release_version}" \
    -srcfolder "${dmg_stage}" \
    -ov \
    -format UDZO \
    "${dmg_path}"

# Sparkle downloads a signed ZIP rather than mounting the user-facing DMG.
update_zip="${output_root}/ReturnBot-${release_version}-arm64.zip"
ditto -c -k --sequesterRsrc --keepParent "${app_path}" "${update_zip}"
print "Update ZIP: ${update_zip}"
print "App: ${app_path}"
print "DMG: ${dmg_path}"
