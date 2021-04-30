import random;

# inp = input("Europe Floor?")
# usf = int(inp) + 1
# print("US Floor ", usf)

slangDictionary = {"one":"cheerio", "two":"pip pip", "three":"smashing"}

slang = ['cheerio', 'pip pip', 'smashing'];
slang.append('cheeky');
slang.append('yonks');
print(slang);
slang.append('delete this slang');
print(slang);
del slang[5];
slang.append('delete this slang');
print(slang);
slang.remove('delete this slang');
print(slang);
result = slangDictionary.get('three');
if result:
    print(result);
else:
    print("Value doesn't exist.");

#random
r = random.random();
print(r);

#range
for i in range(2000, 2020, 2):
    print(i);


#loop and lists
menu = ['knackered spam', 'pip pip spam', 'squidgy spam', 'smashing spam'];
menu_prices = {};
price = 0.50;
for item in menu:
    menu_prices[item] = price;
    price = price+1;

print(menu_prices);

#formatted printing
for name, price in menu_prices.items():
    print(name, ': $', format(price, '.2f'), sep='');