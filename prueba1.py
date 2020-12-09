# Programa imprime los primeros n números primos
print("Ingrese el valor de n mayor a 2\n")

n= int(input())

if (n<=2):
    print("Entrada no válida")
else:
    for i in range(2,n+1):
        tieneDivisores = False
        for j in range(2,i+1):
            if(j==i and not tieneDivisores):
                print(i)
            if(i%j ==0 ):
                tieneDivisores = True

print("fin de programa")
            


