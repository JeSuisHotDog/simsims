import os
from sphinx.application import Sphinx

# Main arguments 
rootdir = os.path.curdir
confdir = os.path.join(rootdir, "conf")
srcdir = os.path.join(rootdir, "conf")
doctreedir = os.path.join(rootdir, "docs")
builddir = os.path.join(doctreedir, "html")
builder = "html"

if __name__ == '__main__':
    # Create the Sphinx application object
    app = Sphinx(srcdir, confdir, builddir, doctreedir, builder)
    # Run the build
    app.build()