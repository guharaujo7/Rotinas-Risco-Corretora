def calcular_resultado(tipo, RMKT, SDP, SPVD, SFD, SPDA, SPTA):
    if tipo == 1:
        resultado = max(0.35 * max(RMKT, 0.25 * SDP, 0.25 * SPVD), SFD)
    elif tipo == 2:
        resultado = max(RMKT, 0.25 * SDP, SFD, 0.18 * SPDA, 0.25 * SPTA, 0.25 * SPVD)
    else:
        raise ValueError("Tipo inválido. Escolha 1 para execução, 2 para liquidação, ou 3 para capacidade econômica.")
    
    return resultado


tipo_calculo = int(input("Escolha o tipo de cálculo (1 para execução, 2 para liquidação, 3 para capacidade econômica): "))

if tipo_calculo == 3:
    
    RMKT = float(input("Insira o valor inicial de RMKT: "))
    SDP = float(input("Insira o valor inicial de SDP: "))
    SPVD = float(input("Insira o valor inicial de SPVD: "))
    SFD = float(input("Insira o valor inicial de SFD: "))
    
    resultado = calcular_resultado(1, RMKT, SDP, SPVD, SFD, 0, 0)
    capacidade_economica = resultado - 25
    
    print(f"A capacidade econômica calculada é: {capacidade_economica:.2f}")
else:
   
    capacidade_economica_desejada = float(input("Insira a capacidade econômica desejada: "))

    
    resultado_desejado = capacidade_economica_desejada + 25

    
    RMKT = float(input("Insira o valor inicial de RMKT: "))
    SDP = float(input("Insira o valor inicial de SDP: "))
    SPVD = float(input("Insira o valor inicial de SPVD: "))
    SFD = float(input("Insira o valor inicial de SFD: "))

   
    if tipo_calculo == 2:
        SPDA = float(input("Insira o valor inicial de SPDA: "))
        SPTA = float(input("Insira o valor inicial de SPTA: "))
    else:
        SPDA = SPTA = 0  

    
    tolerancia = 0.01

    while True:
        resultado_atual = calcular_resultado(tipo_calculo, RMKT, SDP, SPVD, SFD, SPDA, SPTA)
        
        diferenca = resultado_atual - resultado_desejado
        
        if abs(diferenca) <= tolerancia:
            break
        
       
        correcao = 0.1 * diferenca
        max_termo = max(RMKT, 0.25 * SDP, SFD, 0.18 * SPDA, 0.25 * SPTA, 0.25 * SPVD)
        
        if max_termo == RMKT:
            RMKT -= correcao
        elif max_termo == 0.25 * SDP:
            SDP -= 0.25 * correcao
        elif max_termo == 0.25 * SPVD:
            SPVD -= 0.25 * correcao
        elif max_termo == SFD:
            SFD -= correcao
        elif max_termo == 0.18 * SPDA:
            SPDA -= 0.18 * correcao
        elif max_termo == 0.25 * SPTA:
            SPTA -= 0.25 * correcao

   
    print(f"RMKT ajustado: {int(RMKT)}")
    print(f"SDP ajustado: {int(SDP)}")
    print(f"SPVD ajustado: {int(SPVD)}")
    print(f"SFD ajustado: {int(SFD)}")

    if tipo_calculo == 2:
        print(f"SPDA ajustado: {int(SPDA)}")
        print(f"SPTA ajustado: {int(SPTA)}")

