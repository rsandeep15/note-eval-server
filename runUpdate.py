#!/usr/bin/env python
# -*- coding: UTF-8 -*-

# enable debugging
import cgitb, cgi,sys,json
from populatePrice import RealEstate
cgitb.enable()

data = sys.stdin.read()

#
print "Content-Type: text/html;charset=utf-8"
print

request = json.loads(data)

file = request["file"]
user = request["userUid"]

instance = RealEstate()
instance.run(file, user)
