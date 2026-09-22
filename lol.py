from seniatOCR import SeniatOCR


def main():

    variable = SeniatOCR()
    resultado = variable.consultar_con_reintentos("J-08521954-4",12)
    
    print(resultado)

    

if __name__ == "__main__":
    main()