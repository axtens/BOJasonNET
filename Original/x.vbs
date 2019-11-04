set json = createobject("protium.json")

set dict = json.parse(	"{'TITLE' : 'Landing Page 2 - Oh yeah','META' : 'the landing page meta description','BODY' : '<b>This is body text</b>'}")

wscript.echo dict("BODY")
