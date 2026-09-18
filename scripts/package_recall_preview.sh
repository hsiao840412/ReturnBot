#!/bin/zsh
# Build an isolated preview without replacing the existing ReturnBot app or DMG.
set -euo pipefail
project_root="${0:A:h:h}"
cd "$project_root"
export DEVELOPER_DIR="${DEVELOPER_DIR:-/Applications/Xcode.app/Contents/Developer}"
export SWIFTPM_MODULECACHE_OVERRIDE="/tmp/returnbot-recall-swift-cache"
export CLANG_MODULE_CACHE_PATH="/tmp/returnbot-recall-clang-cache"
swift build --disable-sandbox
.venv/bin/python -m PyInstaller --noconfirm \
    --workpath build/recall-helper-work --distpath build/recall-helper-dist ReturnBotHelper.spec
.venv/bin/python -c '
import plistlib, shutil
from pathlib import Path
app = Path("build/recall-preview/ReturnBotRecallPreview-3.2.app")
if app.exists():
    shutil.rmtree(app)
(app / "Contents/MacOS").mkdir(parents=True)
(app / "Contents/Resources").mkdir()
shutil.copy2(".build/debug/ReturnBotMac", app / "Contents/MacOS/ReturnBot")
shutil.copytree("build/recall-helper-dist/ReturnBotHelper", app / "Contents/Resources/ReturnBotHelper")
shutil.copy2("MyIcon.icns", app / "Contents/Resources/MyIcon.icns")
info = plistlib.loads(Path("packaging/Info.plist").read_bytes())
info.update(CFBundleIdentifier="com.returnbot.recall-preview.v32", CFBundleDisplayName="ReturnBot", CFBundleName="ReturnBot", CFBundleShortVersionString="3.2", CFBundleVersion="3.2", RecallPreview=False)
# Preview has a separate identity, so never offer production updates.
info.pop("SUPublicEDKey", None)
info.pop("SUFeedURL", None)
shutil.copytree(".build/artifacts/sparkle/Sparkle/Sparkle.xcframework/macos-arm64_x86_64/Sparkle.framework", app / "Contents/Frameworks/Sparkle.framework", symlinks=True)
(app / "Contents/Info.plist").write_bytes(plistlib.dumps(info))
'
codesign --force --sign - --entitlements packaging/ReturnBot.entitlements \
    build/recall-preview/ReturnBotRecallPreview-3.2.app/Contents/Resources/ReturnBotHelper/ReturnBotHelper
codesign --force --sign - --entitlements packaging/ReturnBot.entitlements \
    build/recall-preview/ReturnBotRecallPreview-3.2.app
codesign --verify --deep --strict build/recall-preview/ReturnBotRecallPreview-3.2.app
print "Preview: ${project_root}/build/recall-preview/ReturnBotRecallPreview-3.2.app"
