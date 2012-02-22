import os
import re

if __name__ == '__main__':
	# Build JumplistZ.exe
	with open('JumplistZ.cpp') as f:
		txt = f.read()
		name = re.findall(r'\s+NAME\[\]\s*=\s*_T\("(.*?)"', txt)
		ver = re.findall(r'\s+VERSION\[\]\s*=\s*_T\("(.*?)"', txt)
		fnout = '%s-%s.exe' % (name[0], ver[0])
	sdkenv = r'C:\Program Files\Microsoft SDKs\Windows\v7.1\Bin\SetEnv.Cmd'
	cmds = [
		'call "%s" /release /x86 /xp' % sdkenv,
		'rc JumplistZ.rc',
		'cl JumplistZ.cpp',
		'link JumplistZ.obj JumplistZ.res /out:%s' % fnout,
		]
	ret = os.system(' && '.join(cmds))
	print 'ret: %d' % ret
