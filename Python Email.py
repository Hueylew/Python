import glob
import emlx

fname = glob.glob('./mails/**/*.emlx', recursive = True)
msg = emlx.read(fname)
print(msg.headers['Subject'])
print('Hello')