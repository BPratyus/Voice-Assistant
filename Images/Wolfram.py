import wolframalpha

client = wolframalpha.Client("A682U3-UE39885KWY")
res = client.query("Who is a cricketer")
try:
    print(res)
    print (next(res.results).text)
    # self.speak (next(res.results).text)
except StopIteration:
    print ("No results")