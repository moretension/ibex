from __future__ import print_function, with_statement

import os
import string
import sys
from shutil import copy2
import zipfile


__all__ = ['IbexError', 'IbexBook', 'Ibex']

    
try:
    # support for Apple's bplist format first appears in CPython 3.4.
    # prior versions only support xml plists.
    from plistlib import FMT_BINARY
except ImportError:
    FMT_BINARY = None

    try:
        # maybe PyObjC is available?
        from Foundation import NSArray, NSDictionary, NSString
    except ImportError:
        NSDictionary = None
        # no plistlib support, no PyObjC, punt to plutil for conversion.
        from plistlib import readPlist
else:
    from plistlib import readPlist


class IbexError(Exception):
    pass

class IbexBook(object):

    """A representation of a book as contained in the iBooks Library.

       The book's metadata may be retrieved using dict conventions, though
       internally the book may be a PyObjC object.
    """
    def __init__(self, ns_book_d, debug=False):
        self.book_d = ns_book_d
        self.debug = debug

    def __getattr__(self, name):
        attr = self.book_d[name]
        
        if not NSDictionary is None:
            if isinstance(attr, NSString):
                return attr.UTF8String()

        return attr

    def export(self, export_dir):
        """Export book to export_dir.

           Epub books will be zipped. Other formats are currently copied
           as found in the iBooks directory structure to export_dir.
        """
        export_target = os.path.join(export_dir, self.BKDisplayName)

        if self.debug:
            print("Exporting to %s" % (export_target,), file=sys.stderr)

        if os.path.isdir(self.path) and self.BKDisplayName.endswith('.epub'):
            cwd = os.getcwd()
            os.chdir(self.path)
            try:
                os_walker = os.walk('.')
                (root, subdirs, files) = next(os_walker)
                if not 'META-INF' in subdirs:
                    # XXX still create zipfile?
                    print("%s: missing META-INF subdirectory in %s" %
                                (self.BKDisplayName, self.path),
                                file=sys.stderr)
                    
                with zipfile.ZipFile(export_target, mode='w',
                                        compression=zipfile.ZIP_STORED) as zf:
                    try:
                        # mimetype must be first entry, uncompressed
                        mt_idx = files.index('mimetype')
                        files.pop(mt_idx)
                        zf.write('mimetype')
                    except ValueError as ve:
                        print("%s: missing mimetype file" % (self.path,),
                                        file=sys.stderr)

                    # remove iTunesMetadata.plist, iBooks will generate it
                    try:
                        itmd_idx = files.index('iTunesMetadata.plist')
                        files.pop(itmd_idx)
                    except:
                        pass

                    # write remaining files at directory depth 0...
                    for f in files:
                        zf.write(f, compress_type=zipfile.ZIP_DEFLATED)

                    # the rest of the hierarchy can be added in walk order
                    for (root, subdirs, files) in os_walker:
                        for f in files:
                            zf.write(os.path.join(root, f),
                                            compress_type=zipfile.ZIP_DEFLATED)

            except zipfile.BadZipfile as bzfe:
                print("%s: %s" % (export_target, str(bzfe)), file=sys.stderr)
            except zipfile.LargeZipFile as lzfe:
                print("%s: %s" % (export_target, str(lzfe)), file=sys.stderr)
            except Exception as e:
                print("%s: %s" % (self.path, str(e)), file=sys.stderr)

            finally:
                os.chdir(cwd)
        else:
            copy2(self.path, export_target)


class Ibex(object):

    """Simple wrapper class for iBooks Library information and export.

       Ibex will try to read the Books.plist using plistlib's bplist support,
       if available, then falls back to PyObjC (NSDictionary). As a last
       resort, attempts to execute plutil(1) manually to convert the bplist to
       xml, reading the converted plist data from stdout.
    """
    PLUTIL_PATH = '/usr/bin/plutil'
    BOOKS_PLIST_PATH = os.path.expanduser('~/Library/Containers/com.apple.BKAgentService/Data/Documents/iBooks/Books/Books.plist')

    @classmethod
    def _ibex_plutil_read_xml(cls, plist_path, plutil_path=PLUTIL_PATH):
        """ use Apple's plutil(1) to convert bplist to xml, dump to stdout """
        from subprocess import Popen, PIPE
        
        # "-o -" means dump converted plist to stdout
        plutil_cmd = [plutil_path, '-convert', 'xml1', '-o', '-', plist_path]
        
        p = Popen(plutil_cmd, bufsize=1, stdout=PIPE, close_fds=True)
        plist = readPlist(p.stdout)
        p.terminate()

        return plist

    def __init__(self, plist_path=BOOKS_PLIST_PATH):
        if FMT_BINARY is None:
            if not NSDictionary is None:
                self.plist = NSDictionary.alloc().initWithContentsOfFile_(plist_path)
            else:
                self.plist = Ibex._ibex_plutil_read_xml(plist_path)
        else:
            self.plist = readPlist(plist_path)

        if self.plist is None:
            raise IbexError('%s: failed to read property list' % plist_path)

    def books(self):
        """Return an iterator of books found in the iBooks Library.

           Yields an IbexBook instance at the current position of the iterator.
        """
        for book in self.plist['Books']:
            yield IbexBook(book, debug=True)

    def export(self, export_path):
        """Export all books found in the iBooks Library to export_path."""
        os.mkdir(os.path.expanduser(export_path), 0o0700)

        for book in self.books():
            book.export(export_path)

    def __del__(self):
        if not NSDictionary is None and not self.plist is None:
            self.plist.release()
            
def main(argv):
    ibex = Ibex(argv[0])
    ibex.export(argv[1])

if __name__ == '__main__':
    SCRIPT = sys.argv[0]
    main(sys.argv[1:])
