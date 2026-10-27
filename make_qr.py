"""Make QR code images.  Usage:
   pip install "qrcode[pil]"
   python make_qr.py https://YOURNAME.github.io/guest-register/ 101 102 103 201 202
Creates qr_main.png (for reception / gate) and qr_room_<no>.png (one per room, opens the form with the room filled in)."""
import sys, qrcode
if len(sys.argv) < 2: sys.exit(__doc__)
base = sys.argv[1].rstrip('/') + '/'
qrcode.make(base).save('qr_main.png')
for room in sys.argv[2:]:
    qrcode.make(f'{base}?room={room}').save(f'qr_room_{room}.png')
print('Done:', 'qr_main.png', *[f'qr_room_{r}.png' for r in sys.argv[2:]])
