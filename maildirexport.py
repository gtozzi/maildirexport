#!/usr/bin/env python3


'''
Exports multiple maildirs into filesystem as EML files
'''

import os
import time
import codecs
import shutil
import mailbox
import logging
import datetime
import email.header
import email.utils
import dateutil.parser


class Main:

	def __init__(self):
		self.log = logging.getLogger('main')

	def run(self, srcPath, dstPath, force=False):
		''' Replicates the Maildir structure into a destinatin EML structure

		@return True on success
		'''

		if not os.path.isdir(srcPath):
			self.log.critical('Source path "%s" is not a directory', srcPath)
			return False

		if not os.path.isdir(dstPath):
			self.log.critical('Destination path "%s" is not a directory', dstPath)
			return False

		dstDir = os.path.join(dstPath, os.path.basename(srcPath))
		if os.path.exists(dstDir):
			if force:
				self.log.warning('Destination dir "%s" already exists, deleting it', dstDir)
				shutil.rmtree(dstDir)
			else:
				self.log.critical('Destination dir "%s" already exists', dstDir)
				return False

		self.log.debug('Creating "%s"', dstDir)
		os.mkdir(dstDir)

		self.exportDir(srcPath, dstDir)

	def exportDir(self, srcPath, dstPath):
		for fileName in os.listdir(srcPath):
			filePath = os.path.join(srcPath, fileName)

			if not os.path.isdir(filePath):
				continue

			if fileName == 'Maildir':
				# Found a maildir, parsing it
				self.log.info('Processing maildir "%s" into "%s"', filePath, dstPath)
				MaildirExporter(filePath, dstPath).export()
				continue

			dstDir = os.path.join(dstPath, fileName)
			self.log.info('Recreating directory "%s" as "%s"', filePath, dstDir)
			os.mkdir(dstDir)
			self.exportDir(filePath, dstDir)


class MaildirExporter:
	''' Exports a single maildir '''

	REPLACECHARS = [ '*', '.',  '"',  '/', '\\',  '[',  ']', ':', ';', '|' ]
	MAXSUBJLEN = 100

	def __init__(self, srcPath, dstPath):
		self.log = logging.getLogger('me')

		self.srcPath = srcPath
		self.dstPath = dstPath

		self.removeChars = [ chr(i) for i in range(0, 32) ]

	def export(self):
		''' Perform the export '''
		self.recursiveExport(self.srcPath, self.dstPath)

	def recursiveExport(self, srcPath, dstPath):
		self.log.debug('Processing folder "%s" into "%s"', srcPath, dstPath)

		assert os.path.isdir(srcPath)
		assert os.path.exists(dstPath)

		md = mailbox.Maildir(srcPath, create=False)

		for message in md:
			if message['Subject']:
				subject = email.header.decode_header(message['Subject'])[0]
				if type(subject[0]) == str:
					subject = subject[0]
				elif type(subject[0]) == bytes:
					if not subject[1] or subject[1].startswith('unknown'):
						subject = subject[0].decode('ascii', errors='replace')
					else:
						try:
							codecs.lookup(subject[1])
						except LookupError:
							subject = subject[0].decode(errors='replace')
						else:
							subject = subject[0].decode(subject[1])
				else:
					raise NotImplementedError(type(subject[0]))
			else:
				subject = '(no subject)'

			date = email.utils.parsedate(message['Date'])
			if date:
				date = datetime.datetime.fromtimestamp(time.mktime(date))
			else:
				date = datetime.datetime.fromtimestamp(0)

			safeSubject = subject.translate({ord(x): '_' for x in self.REPLACECHARS})
			safeSubject = safeSubject.translate({ord(x): '' for x in self.removeChars})
			if len(safeSubject) > self.MAXSUBJLEN:
				safeSubject = safeSubject[:self.MAXSUBJLEN] + '…'

			fileName = "{} {}.eml".format(date.strftime(r'%Y-%m-%d %H:%M:%S'), safeSubject)

			with open(os.path.join(dstPath, fileName), 'wb') as f:
				f.write(bytes(message))

		for folder in md.list_folders():
			subSrcPath = os.path.join(srcPath, '.' + folder)
			subDstPath = os.path.join(dstPath, folder)
			os.mkdir(subDstPath)
			self.recursiveExport(subSrcPath, subDstPath)


if __name__ == '__main__':
	import argparse

	parser = argparse.ArgumentParser(description='Exports multiple maildirs into filesystem as EML files')
	parser.add_argument('source', help='source root path')
	parser.add_argument('dest', help='destinatin root path')
	parser.add_argument('-f', '--force', action='store_true', help='Delete destination dir if existing')

	args = parser.parse_args()

	logging.basicConfig(level=logging.DEBUG)

	Main().run(args.source, args.dest, args.force)

