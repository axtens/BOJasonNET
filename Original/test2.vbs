set json = createobject("protium.json")
set dict = createobject("scripting.dictionary")
dict.add "foo", now()
dict.add "bar", atn(1)*4
wscript.echo json.tostring(dict)
a = array("foo",now(),"bar",atn(1)*4)
wscript.echo json.tostring(a)
a = array("foo",array(now()),"bar",array(atn(1)*4))
wscript.echo json.tostring(a)

