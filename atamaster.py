
import os
import shutil
import flet as ft
from datetime import datetime
from sqlalchemy import create_engine, Column, Integer, String, Boolean, ForeignKey, DateTime, Table
from sqlalchemy.orm import declarative_base, sessionmaker, relationship
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table as RLTable, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors as rl_colors
from pypdf import PdfWriter
import openpyxl

# --- DATABASE ENGINE & SESSION ---
DB_PATH = "atamaster_pro.db"
engine = create_engine(f"sqlite:///{DB_PATH}", connect_args={"check_same_thread": False})
Session = sessionmaker(bind=engine)
Base = declarative_base()

# Junction Tables
meeting_participants = Table(
    'meeting_participants', Base.metadata,
    Column('meeting_id', Integer, ForeignKey('meetings.id'), primary_key=True),
    Column('participant_id', Integer, ForeignKey('participants.id'), primary_key=True)
)

meeting_tasks = Table(
    'meeting_tasks', Base.metadata,
    Column('meeting_id', Integer, ForeignKey('meetings.id'), primary_key=True),
    Column('task_id', Integer, ForeignKey('tasks.id'), primary_key=True)
)

# --- MODELS ---
class Group(Base):
    __tablename__ = 'groups'
    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True, nullable=False)
    description = Column(String)
    meetings = relationship("Meeting", back_populates="group")

class Participant(Base):
    __tablename__ = 'participants'
    id = Column(Integer, primary_key=True)
    name = Column(String, nullable=False)
    email = Column(String)
    company = Column(String)
    tasks = relationship("Task", back_populates="responsible")

class Task(Base):
    __tablename__ = 'tasks'
    id = Column(Integer, primary_key=True)
    description = Column(String, nullable=False)
    status = Column(String, default="OPEN") # OPEN, CLOSED
    participant_id = Column(Integer, ForeignKey('participants.id'))
    deadline_1 = Column(DateTime)
    deadline_2 = Column(DateTime)
    deadline_3 = Column(DateTime)
    responsible = relationship("Participant", back_populates="tasks")
    meetings = relationship("Meeting", secondary=meeting_tasks, back_populates="tasks")

class Meeting(Base):
    __tablename__ = 'meetings'
    id = Column(Integer, primary_key=True)
    title = Column(String, nullable=False)
    date = Column(DateTime, default=datetime.now)
    group_id = Column(Integer, ForeignKey('groups.id'))
    pdf_path = Column(String)
    group = relationship("Group", back_populates="meetings")
    participants = relationship("Participant", secondary=meeting_participants)
    tasks = relationship("Task", secondary=meeting_tasks, back_populates="meetings")

Base.metadata.create_all(engine)

# --- LOGIC CONTROLLER ---
class AtaController:
    def __init__(self):
        self.session = Session()

    def add_group(self, name, description):
        g = Group(name=name, description=description)
        self.session.add(g); self.session.commit()

    def get_groups(self):
        return self.session.query(Group).all()

    def add_participant(self, name, email, company):
        p = Participant(name=name, email=email, company=company)
        self.session.add(p); self.session.commit()

    def get_participants(self):
        return self.session.query(Participant).all()

    def get_open_tasks_for_group(self, group_id):
        if not group_id: return []
        return self.session.query(Task).join(Task.meetings).filter(
            Meeting.group_id == group_id,
            Task.status == "OPEN"
        ).distinct().all()

    def create_meeting(self, title, group_id, task_data, participant_ids):
        meeting = Meeting(title=title, group_id=int(group_id))
        self.session.add(meeting)

        # Add Participants
        participants = self.session.query(Participant).filter(Participant.id.in_(participant_ids)).all()
        meeting.participants = participants

        # Process Tasks
        for t_info in task_data:
            if t_info.get('existing_id'):
                task = self.session.query(Task).get(t_info['existing_id'])
            else:
                task = Task(
                    description=t_info['desc'],
                    participant_id=t_info['resp_id'],
                    deadline_1=t_info['d1'],
                    deadline_2=t_info['d2'],
                    deadline_3=t_info['d3']
                )
                self.session.add(task)
            if task not in meeting.tasks:
                meeting.tasks.append(task)

        self.session.commit()
        return meeting

    def close_task(self, task_id):
        task = self.session.query(Task).get(task_id)
        if task:
            task.status = "CLOSED"
            self.session.commit()

    def get_dashboard_stats(self):
        total_pending = self.session.query(Task).filter(Task.status == "OPEN").count()
        critical = self.session.query(Task).filter(
            Task.status == "OPEN",
            Task.deadline_3 < datetime.now()
        ).count()
        return total_pending, critical

    def get_meetings(self):
        return self.session.query(Meeting).order_by(Meeting.date.desc()).all()

# --- APP UI ---
def main(page: ft.Page):
    page.title = "AtaMaster Pro - Corporate Intelligence"
    page.theme_mode = ft.ThemeMode.DARK
    page.padding = 0

    controller = AtaController()

    # Pickers (Initialize early to avoid 'Unknown Control')
    file_picker = ft.FilePicker()
    attach_picker = ft.FilePicker()
    date_picker = ft.DatePicker()
    page.overlay.extend([file_picker, attach_picker, date_picker])
    page.update()

    # Shared State
    state = {
        'new_meeting_tasks': [],
        'active_deadlines': [None, None, None]
    }

    # --- UI HELPERS ---
    def notify(msg, success=True):
        page.snack_bar = ft.SnackBar(ft.Text(msg), bgcolor="green" if success else "red800")
        page.snack_bar.open = True
        page.update()

    # --- COMPONENTS ---

    def get_dashboard():
        pending, critical = controller.get_dashboard_stats()
        return ft.Column([
            ft.Text("Dashboard Executivo", size=32, weight="bold", color="blue400"),
            ft.Text("Sistema de Gestão de Atas e Pendências", size=16, color="grey"),
            ft.Divider(height=20, color="transparent"),
            ft.Row([
                ft.Container(
                    content=ft.Column([
                        ft.Icon(ft.icons.PENDING_ACTIONS, color="blue200", size=30),
                        ft.Text("TOTAL PENDENTE", size=12, color="blue100", weight="bold"),
                        ft.Text(str(pending), size=40, weight="bold")
                    ], horizontal_alignment="center"),
                    bgcolor="#1e293b", padding=25, border_radius=15, expand=True
                ),
                ft.Container(
                    content=ft.Column([
                        ft.Icon(ft.icons.WARNING_AMBER, color="red200", size=30),
                        ft.Text("CRÍTICO (P3)", size=12, color="red100", weight="bold"),
                        ft.Text(str(critical), size=40, weight="bold", color="red400")
                    ], horizontal_alignment="center"),
                    bgcolor="#1e293b", padding=25, border_radius=15, expand=True,
                    border=ft.border.all(1, "red700") if critical > 0 else None
                )
            ], spacing=15),
            ft.Divider(height=50, color="transparent"),
            ft.Text("Atalhos", size=20, weight="bold"),
            ft.Row([
                ft.FilledButton("Nova Ata", icon=ft.icons.POST_ADD, on_click=lambda _: set_tab(3)),
                ft.FilledButton("Pessoas & Grupos", icon=ft.icons.PEOPLE_OUTLINE, on_click=lambda _: set_tab(2)),
            ])
        ], scroll=ft.ScrollMode.AUTO, expand=True)

    def get_history():
        meetings = controller.get_meetings()
        m_list = ft.Column(spacing=10, scroll=ft.ScrollMode.AUTO)

        def show_details(meeting):
            detail_ui = ft.Column([
                ft.Row([
                    ft.IconButton(ft.icons.ARROW_BACK, on_click=lambda _: set_tab(1)),
                    ft.Text(meeting.title, size=25, weight="bold")
                ]),
                ft.Text(f"Data: {meeting.date.strftime('%d/%m/%Y')} | Grupo: {meeting.group.name}", color="grey"),
                ft.Divider(),
                ft.Text("Itens da Reunião", size=18, weight="bold"),
            ])

            for t in meeting.tasks:
                is_closed = t.status == "CLOSED"
                detail_ui.controls.append(
                    ft.Container(
                        content=ft.Row([
                            ft.Checkbox(value=is_closed, label=t.description,
                                       on_change=lambda e, tid=t.id: toggle_task(tid, e.data)),
                            ft.Container(expand=True),
                            ft.Text(t.responsible.name, size=12, color="grey")
                        ]),
                        padding=10, bgcolor="#0f172a", border_radius=8
                    )
                )

            def toggle_task(tid, val):
                if val == "true":
                    controller.close_task(tid)
                    notify("Item concluído")
                # Re-render details isn't easy here without complex state,
                # but the change is saved.

            detail_ui.controls.append(ft.Divider())
            if meeting.pdf_path:
                detail_ui.controls.append(
                    ft.FilledButton("Ver PDF Original", icon=ft.icons.PICTURE_AS_PDF,
                                   on_click=lambda _: os.startfile(meeting.pdf_path))
                )

            content_area.content = detail_ui
            page.update()

        for m in meetings:
            m_list.controls.append(
                ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.icons.DESCRIPTION, color="blue400"),
                        ft.Column([
                            ft.Text(m.title, weight="bold", size=16),
                            ft.Text(f"{m.date.strftime('%d/%m/%Y')} • {m.group.name if m.group else 'Geral'}", size=12, color="grey")
                        ], expand=True),
                        ft.IconButton(ft.icons.SEARCH, on_click=lambda _, meeting=m: show_details(meeting))
                    ]),
                    padding=15, bgcolor="#1e293b", border_radius=12
                )
            )

        return ft.Column([
            ft.Text("Histórico de Reuniões", size=32, weight="bold"),
            ft.Divider(height=20, color="transparent"),
            m_list
        ], expand=True)

    def get_management():
        groups = controller.get_groups()
        participants = controller.get_participants()

        # Form Group
        gn_in = ft.TextField(label="Nome do Grupo", expand=True)
        gd_in = ft.TextField(label="Descrição", expand=True)
        def save_g(_):
            if gn_in.value:
                controller.add_group(gn_in.value, gd_in.value)
                notify("Grupo criado")
                set_tab(2)

        # Form Person
        pn_in = ft.TextField(label="Nome Completo", expand=True)
        pe_in = ft.TextField(label="E-mail", expand=True)
        pc_in = ft.TextField(label="Empresa / Departamento", expand=True)
        def save_p(_):
            if pn_in.value:
                controller.add_participant(pn_in.value, pe_in.value, pc_in.value)
                notify("Participante cadastrado")
                set_tab(2)

        return ft.Column([
            ft.Text("Gestão de Dados", size=32, weight="bold"),
            ft.Tabs(
                tabs=[
                    ft.Tab(text="Pessoas", content=ft.Column([
                        ft.Row([pn_in, pe_in, pc_in]),
                        ft.FilledButton("Adicionar Pessoa", on_click=save_p),
                        ft.Divider(),
                        ft.Column([ft.ListTile(title=ft.Text(p.name), subtitle=ft.Text(f"{p.company} • {p.email}")) for p in participants], scroll=ft.ScrollMode.AUTO)
                    ], padding=20)),
                    ft.Tab(text="Grupos de Reunião", content=ft.Column([
                        ft.Row([gn_in, gd_in]),
                        ft.FilledButton("Criar Grupo", on_click=save_g),
                        ft.Divider(),
                        ft.Column([ft.ListTile(title=ft.Text(g.name), subtitle=ft.Text(g.description)) for g in groups], scroll=ft.ScrollMode.AUTO)
                    ], padding=20))
                ], expand=True
            )
        ], expand=True)

    def get_new_meeting():
        groups = controller.get_groups()
        participants = controller.get_participants()

        if not groups:
            return ft.Column([
                ft.Icon(ft.icons.WARNING, color="orange", size=50),
                ft.Text("Cadastre pelo menos um GRUPO antes de criar uma ata.", size=20, weight="bold"),
                ft.FilledButton("Ir para Cadastros", on_click=lambda _: set_tab(2))
            ], horizontal_alignment="center", alignment="center", expand=True)

        title_in = ft.TextField(label="Assunto da Reunião", expand=True, border_color="blue400")
        group_drp = ft.Dropdown(
            label="Selecione o Grupo",
            options=[ft.dropdown.Option(str(g.id), g.name) for g in groups],
            on_change=lambda e: load_viva(e.data),
            width=300
        )

        task_ui = ft.Column()

        def load_viva(gid):
            tasks = controller.get_open_tasks_for_group(int(gid))
            state['new_meeting_tasks'] = []
            for t in tasks:
                state['new_meeting_tasks'].append({
                    'desc': t.description,
                    'resp_id': t.participant_id,
                    'resp_name': t.responsible.name,
                    'd1': t.deadline_1,
                    'd2': t.deadline_2,
                    'd3': t.deadline_3,
                    'existing_id': t.id
                })
            refresh_tasks()

        def refresh_tasks():
            task_ui.controls.clear()
            for i, t in enumerate(state['new_meeting_tasks']):
                is_old = t.get('existing_id') is not None
                is_critical = t['d3'] and t['d3'] < datetime.now()
                task_ui.controls.append(
                    ft.Container(
                        content=ft.Row([
                            ft.Icon(ft.icons.HISTORY if is_old else ft.icons.ADD_CIRCLE, color="blue" if is_old else "green"),
                            ft.Column([
                                ft.Text(t['desc'], weight="bold", color="red400" if is_critical else "white"),
                                ft.Text(f"Responsável: {t['resp_name']}", size=12, color="grey")
                            ], expand=True),
                            ft.Text(t['d3'].strftime("%d/%m/%y") if t['d3'] else "", size=12),
                            ft.IconButton(ft.icons.DELETE_OUTLINE, icon_color="red", on_click=lambda _, idx=i: remove_t(idx))
                        ]),
                        padding=12, bgcolor="#0f172a", border_radius=10,
                        border=ft.border.all(1, "red700") if is_critical else None
                    )
                )
            page.update()

        def remove_t(idx):
            state['new_meeting_tasks'].pop(idx)
            refresh_tasks()

        # Add Task
        new_t_desc = ft.TextField(label="Descrição do item...", expand=True, multiline=True)
        new_t_resp = ft.Dropdown(
            label="Responsável",
            options=[ft.dropdown.Option(str(p.id), p.name) for p in participants],
            width=250
        )

        d_row = ft.Row(spacing=5)
        d_btns = []
        for i in range(3):
            btn = ft.OutlinedButton(f"P{i+1}", on_click=lambda _, idx=i: pick_d(idx))
            d_btns.append(btn)
            d_row.controls.append(btn)

        def pick_d(idx):
            def handle_change(e):
                if e.data:
                    dt = datetime.fromisoformat(e.data.split('T')[0])
                    state['active_deadlines'][idx] = dt
                    d_btns[idx].text = dt.strftime("%d/%m")
                    page.update()
            date_picker.on_change = handle_change
            date_picker.pick_date()

        def add_t(_):
            if not new_t_desc.value or not new_t_resp.value:
                return notify("Preencha descrição e responsável", False)

            resp = next(p for p in participants if str(p.id) == new_t_resp.value)
            state['new_meeting_tasks'].append({
                'desc': new_t_desc.value,
                'resp_id': resp.id,
                'resp_name': resp.name,
                'd1': state['active_deadlines'][0],
                'd2': state['active_deadlines'][1],
                'd3': state['active_deadlines'][2]
            })
            new_t_desc.value = ""
            # Reset deadlines
            state['active_deadlines'] = [None, None, None]
            for b in d_btns: b.text = f"P{d_btns.index(b)+1}"
            refresh_tasks()

        # Presence
        presence_row = ft.Row(wrap=True)
        def refresh_presence():
            presence_row.controls.clear()
            for p in participants:
                presence_row.controls.append(ft.Checkbox(label=p.name, value=True, data=p.id))
        refresh_presence()

        # PDF Fusion
        attached_files = []
        attach_list = ft.Row(wrap=True)

        def pick_attachments(_):
            def on_res(e):
                if e.files:
                    for f in e.files:
                        attached_files.append(f.path)
                        attach_list.controls.append(ft.Chip(label=ft.Text(f.name)))
                    page.update()
            attach_picker.on_result = on_res
            attach_picker.pick_files(allow_multiple=True, allowed_extensions=["pdf"])

        return ft.Column([
            ft.Text("Nova Reunião / Ata Viva", size=32, weight="bold"),
            ft.Row([title_in, group_drp]),
            ft.Divider(),
            ft.Text("Lista de Presença", size=18, weight="bold"),
            presence_row,
            ft.Divider(),
            ft.Text("Adicionar Item à Pauta", size=18, weight="bold"),
            ft.Row([new_t_desc, new_t_resp]),
            ft.Row([d_row, ft.FilledButton("Incluir Item", icon=ft.icons.ADD, on_click=add_t)]),
            ft.Divider(),
            ft.Text("Anexos (PDF Fusion)", size=18, weight="bold"),
            ft.Row([
                ft.ElevatedButton("Selecionar PDFs para Mesclar", icon=ft.icons.ATTACH_FILE, on_click=pick_attachments),
                attach_list
            ]),
            ft.Divider(),
            ft.Text("Pauta e Pendências", size=20, weight="bold"),
            task_ui,
            ft.Divider(height=30, color="transparent"),
            ft.FilledButton("FINALIZAR E GERAR PDF", icon=ft.icons.SAVE,
                           bgcolor="blue600", color="white", height=60, expand=True,
                           on_click=lambda _: generate_ata_final(title_in.value, group_drp.value, presence_row.controls, attached_files))
        ], scroll=ft.ScrollMode.AUTO, expand=True)

    def generate_ata_final(title, gid, presence_ctrls, attached_files):
        if not title or not gid: return notify("Título e Grupo obrigatórios", False)

        p_ids = [c.data for c in presence_ctrls if c.value]
        meeting = controller.create_meeting(title, gid, state['new_meeting_tasks'], p_ids)

        # PDF Generation
        pdf_name = f"Ata_{meeting.id}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        doc = SimpleDocTemplate(pdf_name, pagesize=A4)
        styles = getSampleStyleSheet()
        elements = []

        elements.append(Paragraph(f"ATA DE REUNIÃO: {meeting.title}", styles['Title']))
        elements.append(Paragraph(f"Data: {meeting.date.strftime('%d/%m/%Y')} | Grupo: {meeting.group.name}", styles['Normal']))
        elements.append(Spacer(1, 20))

        elements.append(Paragraph("PARTICIPANTES", styles['Heading2']))
        for p in meeting.participants:
            elements.append(Paragraph(f"- {p.name} ({p.company})", styles['Normal']))
        elements.append(Spacer(1, 20))

        elements.append(Paragraph("ITENS DISCUTIDOS E PENDÊNCIAS", styles['Heading2']))
        data = [["Descrição", "Responsável", "Prazo Limite (P3)", "Status"]]
        for t in state['new_meeting_tasks']:
            d3_str = t['d3'].strftime("%d/%m/%Y") if t['d3'] else "---"
            data.append([t['desc'], t['resp_name'], d3_str, "ABERTO"])

        t_style = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), rl_colors.grey),
            ('GRID', (0,0), (-1,-1), 0.5, rl_colors.black),
            ('FONTSIZE', (0,0), (-1,-1), 10)
        ])
        tbl = RLTable(data, colWidths=[250, 100, 100, 50])
        tbl.setStyle(t_style)
        elements.append(tbl)

        elements.append(Spacer(1, 40))
        elements.append(Paragraph("Assinaturas:", styles['Normal']))
        elements.append(Spacer(1, 30))
        elements.append(Paragraph("_____________________________          _____________________________", styles['Normal']))

        doc.build(elements)

        # MESCLAGEM DE PDFs
        if attached_files:
            try:
                merger = PdfWriter()
                merger.append(pdf_name)
                for f in attached_files:
                    merger.append(f)

                final_pdf = f"Ata_Completa_{meeting.id}.pdf"
                with open(final_pdf, "wb") as f_out:
                    merger.write(f_out)

                os.remove(pdf_name)
                pdf_name = final_pdf
            except Exception as e:
                notify(f"Erro na mesclagem: {e}", False)

        meeting.pdf_path = os.path.abspath(pdf_name)
        controller.session.commit()

        notify("Ata gerada e salva com sucesso!")
        set_tab(1)
        os.startfile(pdf_name)

    # --- LAYOUT ---
    sidebar = ft.NavigationRail(
        selected_index=0,
        label_type=ft.NavigationRailLabelType.ALL,
        min_width=100,
        min_extended_width=200,
        group_alignment=-0.9,
        destinations=[
            ft.NavigationRailDestination(icon=ft.icons.DASHBOARD, label="Dashboard"),
            ft.NavigationRailDestination(icon=ft.icons.HISTORY, label="Histórico"),
            ft.NavigationRailDestination(icon=ft.icons.SETTINGS_INPUT_COMPONENT, label="Cadastros"),
            ft.NavigationRailDestination(icon=ft.icons.ADD_TASK, label="Nova Ata"),
        ],
        on_change=lambda e: set_tab(e.control.selected_index)
    )

    content_area = ft.Container(expand=True, padding=40)

    def set_tab(idx):
        sidebar.selected_index = idx
        if idx == 0: content_area.content = get_dashboard()
        elif idx == 1: content_area.content = get_history()
        elif idx == 2: content_area.content = get_management()
        elif idx == 3: content_area.content = get_new_meeting()
        page.update()

    page.add(
        ft.Row([
            sidebar,
            ft.VerticalDivider(width=1),
            content_area
        ], expand=True)
    )

    set_tab(0)

if __name__ == "__main__":
    ft.app(target=main)
