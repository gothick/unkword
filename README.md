# unkword

This is a Python script that seems to work okay to extract the text content of some elderly KOffice KWord (.kwd)
files I had lying around. They were 2002/2003 vintage. This copes with a couple of different formats. This was
a very quick-and-dirty two-hour project I threw together to use once, on all my old files, so it's entirely
unsupported and the code is the kind of quality you'd expect. But if you want to rescue some old text from some
old KWord files, you might find it handy. You'll probably need to install `lxml`; I think everything else I used
was a Python builtin.

There were a couple of different file formats I encountered: the .kwd files themselves were sometimes really a
zip file, and sometimes really a gzip tarball, and the main document inside (in both cases 'maindoc.xml') also
varied, sometimes with the XML content being namespaced and sometimes not. The embedded XSLT transform document
you'll find in the Python script copes with either.

This is *NOT A CONVERTER*, really, it just extracts the text content from the KWord documents, which was all I
needed. I'm just making it public in case anyone else has some very old KWord files they need to get the basic
words out of.
