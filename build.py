import os
import re

if __name__ == '__main__':
	# JumplistZ.sample.ini --> JumplistZ.sample.h
	with open('JumplistZ.sample.ini') as f:
		with open('JumplistZ.sample.h', 'w') as o:
			o.write('\nLPCSTR szSampleINI =\n');
			for line in f:
				str = re.sub(r'(\\|")', r'\\\g<1>', line.strip())
				newline = '\t"%s\\n"\n' % str.ljust(65, ' ')
				o.write(newline);
			o.write('\t"";\n');

	# Build JumplistZ.exe
	sdkenv = r'C:\Program Files\Microsoft SDKs\Windows\v7.1\Bin\SetEnv.Cmd'
	os.system(r'call "%s" /release /x86 /xp && cl JumplistZ.cpp' % sdkenv);

