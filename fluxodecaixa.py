import openpyxl
from openpyxl import Workbook

#Informações do usuário

def pegar_valor():
    valor = float(input("Qual foi o valor? "))
    return valor

def pegar_tipo():
    tipo = input("Qual foi o tipo de movimentação? (lucro/gasto) ").lower()
    return tipo

def pegar_descricao():
    descricao = input("Qual foi a descrição do lucro ou gasto? ")
    return descricao

def qual_origem():
    origem = input("Qual foi a origem do lucro ou gasto? ")
    return origem

movimentacoes = []

#Funções principais

continuar = input("Deseja adicionar uma movimentação? (S/N) ").upper()

while continuar == "S":
    valor = pegar_valor()
    tipo = pegar_tipo()
    descricao = pegar_descricao()
    origem = qual_origem()
    movimentacao = {
        "valor": valor,
        "tipo": tipo,
        "descricao": descricao,
        "origem": origem
    }

    movimentacoes.append(movimentacao)
    print("Movimentação adicionada com sucesso!")
    continuar = input("Deseja adicionar outra movimentação? (S/N) ").upper()


def calcular_saldo(movimentacoes):
    saldo = 0.0
    for movimentacao in movimentacoes:
        if movimentacao["tipo"].lower() == "lucro":
            saldo += movimentacao["valor"]
        elif movimentacao["tipo"].lower() == "gasto":
            saldo -= movimentacao["valor"]
    return saldo
resultado_saldo = calcular_saldo(movimentacoes)

#Parte visual

mostrar_fluxo_de_caixa = input("Deseja mostrar o fluxo de caixa? (S/N) ").upper()

if mostrar_fluxo_de_caixa == "S":

    print("\nFluxo de Caixa:")
    print("==============================")

    for movimentacao in movimentacoes:

        print(f"Descrição: {movimentacao['descricao']}")
        print(f"Origem: {movimentacao['origem']}")
        print(f"Tipo: {movimentacao['tipo']}")
        print(f"Valor: R$ {movimentacao['valor']:.2f}")

        print("------------------------------")

    print(f"Saldo: R$ {resultado_saldo:.2f}")

def gerar_planilha():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Fluxo de Caixa"
    worksheet["A1"] = "Descrição"
    worksheet["B1"] = "Origem"
    worksheet["C1"] = "Tipo"
    worksheet["D1"] = "Valor"

    for movimentacao in movimentacoes:
        worksheet.append([
            movimentacao["descricao"],
            movimentacao["origem"],
            movimentacao["tipo"],
            movimentacao["valor"]
        ])
    workbook.save("fluxo_de_caixa.xlsx")

gerar_planilha()