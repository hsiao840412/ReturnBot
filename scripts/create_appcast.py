#!/usr/bin/env python3
"""Sign a packaged release with the ReturnBot key in the local macOS Keychain."""
import email.utils
import pathlib
import plistlib
import subprocess
import sys
import xml.etree.ElementTree as ET

root = pathlib.Path(__file__).resolve().parent.parent
version = sys.argv[1]
folder = root / 'build' / f'macos-{version}'
archive = folder / (sys.argv[2] if len(sys.argv) > 2 else f'ReturnBot-{version}-arm64.zip')
assert archive.parent == folder and archive.suffix == '.zip'
info = plistlib.loads((folder / 'ReturnBot.app/Contents/Info.plist').read_bytes())
assert info['CFBundleVersion'] == version
signer = root / '.build/artifacts/sparkle/Sparkle/bin/sign_update'
command = [str(signer), '--account', 'com.returnbot.app']
signature = subprocess.check_output(command + ['-p', str(archive)], text=True).strip()
subprocess.run(command + ['--verify', str(archive), signature], check=True)
ns = 'http://www.andymatuschak.org/xml-namespaces/sparkle'
ET.register_namespace('sparkle', ns)
rss = ET.Element('rss', version='2.0')
channel = ET.SubElement(rss, 'channel')
ET.SubElement(channel, 'title').text = 'ReturnBot 更新'
item = ET.SubElement(channel, 'item')
ET.SubElement(item, 'title').text = f'ReturnBot {version}'
ET.SubElement(item, 'pubDate').text = email.utils.formatdate(usegmt=True)
ET.SubElement(item, f'{{{ns}}}version').text = version
ET.SubElement(item, f'{{{ns}}}shortVersionString').text = version
ET.SubElement(item, f'{{{ns}}}minimumSystemVersion').text = info['LSMinimumSystemVersion']
ET.SubElement(item, 'description').text = '寄銷召回：交換價格為 0 時改用庫存價格，並顯示提示；移除精確辨識選項，保留快速辨識。請重新匯入含庫存價格的價格表。'
ET.SubElement(item, 'enclosure', {
    'url': f'https://github.com/hsiao840412/ReturnBot/releases/download/v{version}/{archive.name}',
    'length': str(archive.stat().st_size), 'type': 'application/octet-stream',
    f'{{{ns}}}edSignature': signature,
})
feed = folder / 'appcast.xml'
ET.indent(rss)
ET.ElementTree(rss).write(feed, encoding='utf-8', xml_declaration=True)
subprocess.run(command + [str(feed)], check=True)
subprocess.run(command + ['--verify', str(feed)], check=True)
print(feed)
