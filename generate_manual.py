from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem
from reportlab.lib.styles import getSampleStyleSheet

def generate_manual():
    doc = SimpleDocTemplate("Manual_AtaMaster.pdf", pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    elements.append(Paragraph("Manual do Usuário: AtaMaster Pro", styles['Title']))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("1. Introdução", styles['Heading1']))
    elements.append(Paragraph("O AtaMaster Pro é uma ferramenta para gestão de atas de reunião focada em tarefas contínuas. O diferencial é a 'Ata Viva', que mantém o rastreio de pendências entre encontros.", styles['Normal']))

    elements.append(Spacer(1, 12))
    elements.append(Paragraph("2. Fluxo de Trabalho", styles['Heading1']))
    items = [
        ListItem(Paragraph("<b>Cadastro:</b> Comece criando Grupos e Participantes na aba 'Gestão'.")),
        ListItem(Paragraph("<b>Nova Reunião:</b> Selecione um grupo. Itens abertos de reuniões anteriores aparecerão automaticamente.")),
        ListItem(Paragraph("<b>Adição de Itens:</b> Insira novas tarefas com até 3 prazos de acompanhamento.")),
        ListItem(Paragraph("<b>Finalização:</b> Gere o PDF para coleta de assinaturas.")),
        ListItem(Paragraph("<b>Acompanhamento:</b> Use o Dashboard para ver alertas de itens atrasados.")),
    ]
    elements.append(ListFlowable(items, bulletType='bullet'))

    elements.append(Spacer(1, 12))
    elements.append(Paragraph("3. Gestão de Prazos", styles['Heading1']))
    elements.append(Paragraph("O sistema permite definir 3 datas. Ao ultrapassar a 3ª data, o item ganha um destaque vermelho crítico no Dashboard.", styles['Normal']))

    elements.append(Spacer(1, 24))
    elements.append(Paragraph("Desenvolvido por Daniel Alves Anversi", styles['Italic']))

    doc.build(elements)
    print("Manual gerado com sucesso: Manual_AtaMaster.pdf")

if __name__ == "__main__":
    generate_manual()
