import os
import string
import sys
from shutil import copy2
import zipfile

from Foundation import NSArray, NSDictionary, NSString


class IbexError(Exception):
    pass

class IbexBook(object):
    def __init__(self, ns_book_d, debug=False):
	self.book_d = ns_book_d
	self.debug = debug

    def __getattr__(self, name):
	attr = self.book_d[name]
	if isinstance(attr, NSString):
	    return attr.UTF8String()
	else:
	    return attr

    def export(self, export_path):
	export_target = os.path.join(export_path, self.BKDisplayName)

	if self.debug:
	    print >>sys.stderr, "Exporting to %s" % (export_target,)

	if os.path.isdir(self.path) and self.BKDisplayName.endswith('.epub'):
	    cwd = os.getcwd()
	    os.chdir(self.path)
	    try:
		os_walker = os.walk('.')
		(root, subdirs, files) = os_walker.next()
		if not 'META-INF' in subdirs:
		    """ XXX still create zipfile? """
		    print >>sys.stderr, \
				"%s: missing META-INF subdirectory in %s" % \
				(self.BKDisplayName, self.path)
		    
		zf = zipfile.ZipFile(export_target, mode='w',
					compression=zipfile.ZIP_STORED)
		try:
		    """ mimetype must be first entry, uncompressed """
		    mt_idx = files.index('mimetype')
		    files.pop(mt_idx)
		    zf.write('mimetype')
		except ValueError, ve:
		    print >>sys.stderr, "%s: missing mimetype file" % \
				(self.path,)

		""" remove iTunesMetadata.plist, iBooks will generate it """
		try:
		    itmd_idx = files.index('iTunesMetadata.plist')
		    files.pop(itmd_idx)
		except:
		    pass

		""" write remaining files at directory depth 0... """
		for f in files:
		    zf.write(f, compress_type=zipfile.ZIP_DEFLATED)

		""" the rest of the hierarchy can be added in walk order """
		for (root, subdirs, files) in os_walker:
		    for f in files:
			zf.write(os.path.join(root, f),
					compress_type=zipfile.ZIP_DEFLATED)

		"""
		for d in subdirs:
		    for (root, dirs, files) in os.walk(d):
			for f in files:
			    zf.write(os.path.join(root, f))
		"""
	    
	    except zipfile.BadZipfile, bzfe:
		print >>sys.stderr, "%s: %s" % (export_target, str(bzfe))
	    except zipfile.LargeZipFile, lzfe:
		print >>sys.stderr, "%s: %s" % (export_target, str(lzfe))
	    except Exception, e:
		print >>sys.stderr, "ERROR: %s: %s" % (self.path, str(e))

	    finally:
		zf.close()
		os.chdir(cwd)
	else:
	    copy2(self.path, export_target)


class Ibex(object):
    BOOKS_PLIST_PATH = os.path.expanduser('~/Library/Containers/com.apple.BKAgentService/Data/Documents/iBooks/Books/Books.plist')

    def __init__(self, plist_path=BOOKS_PLIST_PATH):
	self.plist = NSDictionary.alloc().initWithContentsOfFile_(plist_path)
	if self.plist is None:
	    raise IbexError('%s: failed to read property list' % plist_path)

    def books(self):
	for book in self.plist['Books']:
	    yield IbexBook(book, debug=True)

    def export(self, export_path):
	os.mkdir(os.path.expanduser(export_path), 0700)

	for book in self.books():
	    book.export(export_path)

    def __del__(self):
	if not self.plist is None:
	    self.plist.release()
	    
def main(argv):
    ibex = Ibex(argv[0])
    ibex.export(argv[1])

if __name__ == '__main__':
    SCRIPT = sys.argv[0]
    main(sys.argv[1:])
