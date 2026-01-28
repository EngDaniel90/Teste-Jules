from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

def generate_manual():
    filename = "Manual_AtaMaster.pdf"
    doc = SimpleDocTemplate(filename, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    # Custom Title Style
    title_style = ParagraphStyle(
        'MainTitle',
        parent=styles['Title'],
        fontSize=24,
        spaceAfter=30
    )

    header_style = styles['Heading1']
    sub_header_style = styles['Heading2']
    normal_style = styles['Normal']

    # --- Title Page ---
    elements.append(Spacer(1, 100))
    elements.append(Paragraph("Manual do Usuário: AtaMaster Pro", title_style))
    elements.append(Paragraph("Sistema Moderno de Gestão de Atas de Reunião", styles['Heading2']))
    elements.append(Spacer(1, 50))
    elements.append(Paragraph("Desenvolvido por: Jules (AI Assistant)", normal_style))
    elements.append(Paragraph(f"Data: 28/05/2025", normal_style))
    elements.append(PageBreak())

    # --- 1. Dashboard & Alertas ---
    elements.append(Paragraph("1. Dashboard & Alertas", header_style))
    elements.append(Paragraph(
        "A tela inicial fornece uma visão geral rápida do sistema. Ela exibe cartões informativos com o total de tarefas, "
        "tarefas em aberto e, mais importante, o número de Alertas Críticos.", normal_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(
        "O sistema monitora automaticamente os três prazos definidos para cada tarefa. Quando o terceiro prazo é ultrapassado "
        "e a tarefa continua aberta, ela aparece na lista de 'Alertas Críticos' com destaque em vermelho, garantindo que "
        "nenhum item importante seja esquecido.", normal_style))
    elements.append(Spacer(1, 24))

    # --- 2. Gestão de Grupos e Pessoas ---
    elements.append(Paragraph("2. Gestão de Grupos e Pessoas", header_style))
    elements.append(Paragraph(
        "No menu 'Grupos & Pessoas', você pode organizar suas reuniões por contexto (ex: Engenharia, Projetos, RH).", normal_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(
        "Também é possível cadastrar os participantes fixos, informando nome, email e empresa. Esses dados serão utilizados "
        "para designar responsáveis pelas tarefas durante a criação das atas.", normal_style))
    elements.append(Spacer(1, 24))

    # --- 3. Criação de Atas (Ata Viva) ---
    elements.append(Paragraph("3. Criação de Atas (Ata Viva)", header_style))
    elements.append(Paragraph(
        "Esta é a funcionalidade principal do sistema. Ao iniciar uma 'Nova Reunião', você seleciona o Grupo correspondente.", normal_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(
        "O conceito de 'Ata Viva' significa que todas as tarefas que permaneceram em aberto na reunião anterior do mesmo grupo "
        "são importadas automaticamente para a nova pauta. Isso garante continuidade e acompanhamento contínuo dos itens pendentes.", normal_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(
        "Durante a criação, você pode adicionar novos itens. O sistema gera os bullets automaticamente para facilitar a leitura. "
        "Para cada item, você designa um responsável e define até 3 datas de acompanhamento.", normal_style))
    elements.append(Spacer(1, 24))

    # --- 4. Personalização Visual ---
    elements.append(Paragraph("4. Personalização Visual", header_style))
    elements.append(Paragraph(
        "No menu 'Estilos & Backup', você pode deixar o AtaMaster Pro com a sua cara.", normal_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(
        "- **Modos de Visualização:** Alterne entre o Modo Escuro (Dark) para maior conforto visual ou o Modo Claro (Light).", normal_style))
    elements.append(Paragraph(
        "- **Cores do Sistema:** Escolha entre 3 cores básicas (Azul, Verde, Laranja) para mudar o tom principal do aplicativo.", normal_style))
    elements.append(Spacer(1, 24))

    # --- 5. Backup e Restauração ---
    elements.append(Paragraph("5. Backup e Restauração", header_style))
    elements.append(Paragraph(
        "Para garantir a segurança dos seus dados, o sistema oferece uma funcionalidade de Backup no mesmo menu de Estilos.", normal_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(
        "- **Exportar Backup:** Cria uma cópia de segurança de todo o banco de dados (.db) em um local de sua escolha.", normal_style))
    elements.append(Paragraph(
        "- **Importar Backup:** Permite restaurar todos os seus grupos e atas a partir de um arquivo de backup.", normal_style))
    elements.append(Spacer(1, 24))

    # --- 6. Exportação PDF ---
    elements.append(Paragraph("6. Exportação PDF", header_style))
    elements.append(Paragraph(
        "Ao finalizar uma reunião ou consultar o histórico em 'Minhas Atas', você pode gerar um arquivo PDF profissional. "
        "O PDF contém o cabeçalho com os detalhes da reunião, uma tabela organizada com todas as tarefas e campos de "
        "assinatura para todos os participantes envolvidos.", normal_style))

    doc.build(elements)
    print(f"Manual gerado com sucesso: {filename}")

if __name__ == "__main__":
    generate_manual()
