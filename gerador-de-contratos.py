import json
from datetime import datetime
from docx import Document

 
cursos = []

# Salva a lista de cursos em um arquivo .json
def salvar_cursos(lista_cursos, nome_arquivo = 'cursos.json'):
    with open(nome_arquivo, 'w') as arquivo:
        json.dump(lista_cursos, arquivo, indent = 4)

# Carrega a lista de cursos 
def carregar_cursos(nome_arquivo = 'cursos.json'):
    try:
        with open(nome_arquivo, 'r') as arquivo:
            return json.load(arquivo)
    except (FileNotFoundError, json.JSONDecodeError):
        return []

# Função para registrar cursos    
def registrar_cursos(nome,cargaHoraria, valor):
    curso = {
        'Nome do curso' : nome,
        'Carga horaria do curso' : cargaHoraria,
        'Valor do curso' : valor
    }
    cursos.append(curso)
    salvar_cursos(cursos)

cursos = carregar_cursos()

# Função para remover curso 
def remover_curso():
    if not cursos:
        print("Nenhum curso registrado ainda.")
        return

    print("\nCursos cadastrados:")
    for i, curso in enumerate(cursos, start=1):
        print(f"[{i}] {curso['Nome do curso']} - {curso['Carga horaria do curso']} - R$ {curso['Valor do curso']}")

    try:
        opcao = int(input("Digite o número do curso que deseja remover: "))
        if 1 <= opcao <= len(cursos):
            curso_removido = cursos.pop(opcao - 1)
            salvar_cursos(cursos)
            print(f"\nCurso '{curso_removido['Nome do curso']}' removido com sucesso!\n")
        else:
            print("Número inválido.")
    except ValueError:
        print("Por favor, digite um número válido.")

def formatar_cpf(cpf):
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"

def gerar_contrato():
    if not cursos:
        print("Nenhum curso registrado ainda.")
        return
    
    print("\nCursos cadastrados:")
    for i, curso in enumerate(cursos, start=1):
        print(f"[{i}] {curso['Nome do curso']} - {curso['Carga horaria do curso']} - R$ {curso['Valor do curso']}")

    try:
        escolha = int(input("Digite o número do curso para gerar o contrato: ")) - 1
        curso_escolhido = cursos[escolha]

        # Coletar informações para o contrato 
        nome_aluno = str(input('Digite o nome do aluno: '))
        cpf_aluno = str(input('Digite o cpf do aluno(Sem os pontos): '))
        cpf_aluno = formatar_cpf(cpf_aluno)
        data_de_inicio = str(input('Insira a data de inicio do curso (Use o formato dd/mm/aaaa): '))
        data_de_termino = str(input('Insira a data de termino do curso (Use o formato dd/mm/aaaa): '))
        forma_de_pagamento = str(input('Insira a forma de pagamento: '))
        data_atual = datetime.now().strftime('%d/%m/%Y')

        # Carrega o modelo salvo 
        doc = Document('modeloDeContrato.docx')
        
        # Gerar contrato a partir do modelo salvo 
        for par in doc.paragraphs:
            par.text = par.text.replace('{NOME_DO_ALUNO}' , nome_aluno)
            par.text = par.text.replace('{CPF_DO_ALUNO}' , cpf_aluno)
            par.text = par.text.replace('{NOME_DO_CURSO}' , curso_escolhido['Nome do curso'])
            par.text = par.text.replace('{DURACAO_DO_CURSO}' , curso_escolhido['Carga horaria do curso'])
            par.text = par.text.replace('{DATA_INICIO}' , data_de_inicio)
            par.text = par.text.replace('{DATA_TERMINO}' , data_de_termino)
            par.text = par.text.replace('{VALOR_DO_CURSO}' , str(curso_escolhido['Valor do curso']))
            par.text = par.text.replace('{FORMA_DE_PAGAMENTO}' , forma_de_pagamento)
            par.text = par.text.replace('{DATA_ASSINATURA}' , data_atual)

        nome_arquivo = f"contrato_{nome_aluno.replace(' ', '_')}.docx"
        doc.save(nome_arquivo)
        print(f"Contrato salvo como: {nome_arquivo}")

           

    except (ValueError, IndexError):
        print("Escolha inválida.")
        return
    


def tela_menu():
    print('=-' * 25)
    print('Techrise - Gerador de contratos')
    print('=-' * 25)
    print('')
    print('')
    print('[1] Registrar curso')
    print('[2] Remover curso')
    print('[3] Gerar contrato')
    print('[4] Sair')

while True:
    tela_menu()
    esco = int(input('Digite a opção desejada: '))
    if esco == 1:
        nomeCurso = str(input('Insira o nome do curso: '))
        cargaHoraria = str(input('Insira a carga horaria do curso(Insira apenas os numero): '))
        cargaHoraria = f'{cargaHoraria} Horas'
        valor = float(input('Insira o valor do curso: '))
        registrar_cursos(nomeCurso, cargaHoraria, valor)
        print('Curso adicionado com sucesso!')

    elif esco == 2:
        remover_curso()

    elif esco == 3:
        gerar_contrato()
    
    elif esco == 4:
        break