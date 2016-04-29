# msgtodf
----
Extracts .msg files and outputs to a pandas dataframe. This is a modified version of Matthew Walker's [Extractmsg script](https://github.com/mattgwwalker/msg-extractor), and so functions on the same basis: using Philippe Lagadec's [olefile module](http://www.decalage.info/python/olefileio), that reads Microsoft OLE2 files. 

##Usage
The script takes in a path, as well as an indication of whether the path is a directory(for multiple files) or a single file. The output is a pandas df containing the to, from, cc, date, subject and body from the .msg file.
```python
msg=Extract('path',isdir=True)
mydf = msg.makedf()
```

##Installation
You can install using pip.
```python
 pip install https://github.com/zhilingc/msgtodf/zipball/master
```


