import pandas as pd


def read_excel_file(file_name):
    # usando o context manager 'with' para garantir o fechamento correto do arquivo
    with pd.ExcelFile(file_name) as xls:
        return pd.read_excel(xls)


def convert_grades_to_float(table, columns):
    # conversão direta de int para float dentro do DataFrame existente
    table[columns] = table[columns].astype(float) / 10


def calculate_final_average(table, columns):
    # cálculo da média e adição direta como nova coluna no DataFrame
    table['Média Final'] = table[columns].mean(axis=1).round()


def student_situation(table, approved, naf):
    # simplificação das condições para definir a situação do aluno
    table['Situação'] = 'Aprovado'
    table.loc[table['Média Final'] < naf, 'Situação'] = 'Reprovado'
    table.loc[(table['Média Final'] < approved) & (table['Média Final'] >= naf), 'Situação'] = 'Exame Final'


def failed_attendance(table, missed_classes_limit):
    # lógica para marcar alunos reprovados por falta
    table.loc[table['Faltas'] > missed_classes_limit, 'Situação'] = 'Reprovado por Falta'


def final_exam_calc(table, approved):
    # cálculo da nota para aprovação no exame final
    table['Nota para Aprovação Final'] = 0
    table.loc[table['Situação'] == 'Exame Final', 'Nota para Aprovação Final'] = approved - table['Média Final']


def main():
    # definição de variáveis e parâmetros
    file_name = 'planilha_alunos.xlsx'
    grade_columns = ['P1', 'P2', 'P3']
    approved = 7
    naf = 5
    semester_classes = 60
    missed_classes_limit = 0.25 * semester_classes

    # processamento dos dados seguindo uma sequência lógica e clara
    table = read_excel_file(file_name)
    convert_grades_to_float(table, grade_columns)
    calculate_final_average(table, grade_columns)
    student_situation(table, approved, naf)
    failed_attendance(table, missed_classes_limit)  # Ajuste conforme a correção
    final_exam_calc(table, approved)

    # exibição do DataFrame final
    print('=== Tabela Atualizada ===\n')
    print(table.to_string())


if __name__ == '__main__':
    main()
