import os
import re


# -- Edit this first --
PATH_SETENV = r'C:\Program Files\Microsoft SDKs\Windows\v7.1\Bin\SetEnv.Cmd'


def main():
	# Get new name
	with open('JumplistZ.cpp') as f:
		txt = f.read()
	name = re.findall(r'\s+NAME\[\]\s*=\s*_T\("(.*?)"', txt)
	version = re.findall(r'\s+VERSION\[\]\s*=\s*_T\("(.*?)"', txt)
	fnout = '%s-%s.exe' % (name[0], version[0])

	# Build JumplistZ.exe
	cmds = [
		'call "%s" /release /x86 /xp' % PATH_SETENV,
		'rc JumplistZ.rc',
		'cl JumplistZ.cpp',
		'link JumplistZ.obj JumplistZ.res /out:%s' % fnout,
		]
	ret = os.system(' && '.join(cmds))
	print 'ret: %d' % ret


if __name__ == '__main__':
	main()
