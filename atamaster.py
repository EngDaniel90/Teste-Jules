import os
import shutil
from datetime import datetime
from sqlalchemy import create_engine, Column, Integer, String, Text, DateTime, ForeignKey, Enum, Table
from sqlalchemy.orm import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
import enum
import flet as ft
try:
    import openpyxl
except ImportError:
    openpyxl = None
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table as RelTable, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

Base = declarative_base()

class StatusEnum(enum.Enum):
    OPEN = "OPEN"
    CLOSED = "CLOSED"

# Junction table for Meeting and Task to preserve history
meeting_tasks = Table('meeting_tasks', Base.metadata,
    Column('meeting_id', Integer, ForeignKey('meetings.id'), primary_key=True),
    Column('task_id', Integer, ForeignKey('tasks.id'), primary_key=True)
)

class Group(Base):
    __tablename__ = 'groups'
    id = Column(Integer, primary_key=True)
    name = Column(String(255), nullable=False)
    description = Column(Text)
    created_at = Column(DateTime, default=datetime.utcnow)
    meetings = relationship("Meeting", back_populates="group", cascade="all, delete-orphan")

class Participant(Base):
    __tablename__ = 'participants'
    id = Column(Integer, primary_key=True)
    name = Column(String(255), nullable=False)
    email = Column(String(255))
    company = Column(String(255))
    tasks = relationship("Task", back_populates="responsible")

class Meeting(Base):
    __tablename__ = 'meetings'
    id = Column(Integer, primary_key=True)
    group_id = Column(Integer, ForeignKey('groups.id'))
    title = Column(String(255), nullable=False)
    date = Column(DateTime, default=datetime.utcnow)
    location = Column(String(255))
    group = relationship("Group", back_populates="meetings")
    tasks = relationship("Task", secondary=meeting_tasks, back_populates="meetings")

class Task(Base):
    __tablename__ = 'tasks'
    id = Column(Integer, primary_key=True)
    meeting_origin_id = Column(Integer, ForeignKey('meetings.id'))
    description = Column(Text, nullable=False)
    responsible_id = Column(Integer, ForeignKey('participants.id'))
    status = Column(Enum(StatusEnum), default=StatusEnum.OPEN)
    date1 = Column(DateTime)
    date2 = Column(DateTime)
    date3 = Column(DateTime)
    meetings = relationship("Meeting", secondary=meeting_tasks, back_populates="tasks")
    responsible = relationship("Participant", back_populates="tasks")

DB_URL = "sqlite:///atamaster.db"
engine = create_engine(DB_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

def init_db():
    Base.metadata.create_all(bind=engine)

def get_session():
    return SessionLocal()

# --- CRUD FUNCTIONS ---
def create_group(name, description=None):
    with get_session() as session:
        group = Group(name=name, description=description)
        session.add(group); session.commit(); session.refresh(group)
        return group

def get_groups():
    with get_session() as session: return session.query(Group).all()

def delete_group(group_id):
    with get_session() as session:
        group = session.query(Group).filter(Group.id == group_id).first()
        if group: session.delete(group); session.commit()

def create_participant(name, email=None, company=None):
    with get_session() as session:
        p = Participant(name=name, email=email, company=company)
        session.add(p); session.commit(); session.refresh(p)
        return p

def get_participants():
    with get_session() as session: return session.query(Participant).all()

def delete_participant(p_id):
    with get_session() as session:
        p = session.query(Participant).filter(Participant.id == p_id).first()
        if p: session.delete(p); session.commit()

def create_meeting(group_id, title, date=None, location=None):
    with get_session() as session:
        m = Meeting(group_id=group_id, title=title, date=date or datetime.utcnow(), location=location)
        session.add(m); session.commit(); session.refresh(m)
        return m

def get_meetings_by_group(group_id):
    with get_session() as session:
        return session.query(Meeting).filter(Meeting.group_id == group_id).order_by(Meeting.date.desc()).all()

def create_task(meeting_id, description, responsible_id, date1=None, date2=None, date3=None):
    with get_session() as session:
        task = Task(meeting_origin_id=meeting_id, description=description, responsible_id=responsible_id, date1=date1, date2=date2, date3=date3)
        session.add(task)
        m = session.query(Meeting).filter(Meeting.id == meeting_id).first()
        if m: m.tasks.append(task)
        session.commit(); session.refresh(task)
        return task

def get_tasks_by_meeting(meeting_id):
    with get_session() as session:
        m = session.query(Meeting).filter(Meeting.id == meeting_id).first()
        if not m: return []
        # Accessing m.tasks inside session or before session closes
        return [t for t in m.tasks]

def update_task_status(task_id, status):
    with get_session() as session:
        task = session.query(Task).filter(Task.id == task_id).first()
        if task: task.status = status; session.commit()
        return task

def get_open_tasks_for_group(group_id):
    with get_session() as session:
        return session.query(Task).join(meeting_tasks).join(Meeting)\
            .filter(Meeting.group_id == group_id)\
            .filter(Task.status == StatusEnum.OPEN).distinct().all()

def carry_over_tasks(group_id, new_meeting_id):
    with get_session() as session:
        open_tasks = session.query(Task).join(meeting_tasks).join(Meeting)\
            .filter(Meeting.group_id == group_id)\
            .filter(Task.status == StatusEnum.OPEN).distinct().all()
        m = session.query(Meeting).filter(Meeting.id == new_meeting_id).first()
        if m:
            for task in open_tasks:
                if task not in m.tasks: m.tasks.append(task)
        session.commit()

def generate_pdf(meeting_id):
    with get_session() as session:
        meeting = session.query(Meeting).filter(Meeting.id == meeting_id).first()
        if not meeting: return None
        tasks = meeting.tasks
        filename = f"ata_reuniao_{meeting_id}.pdf"
        doc = SimpleDocTemplate(filename, pagesize=A4)
        elements = []
        styles = getSampleStyleSheet()
        elements.append(Paragraph(f"ATA DE REUNIÃO: {meeting.title}", styles['Title']))
        elements.append(Paragraph(f"Grupo: {meeting.group.name}", styles['Normal']))
        elements.append(Paragraph(f"Data: {meeting.date.strftime('%d/%m/%Y')}", styles['Normal']))
        elements.append(Paragraph(f"Local: {meeting.location or 'N/A'}", styles['Normal']))
        elements.append(Spacer(1, 12))
        data = [["Descrição", "Responsável", "Empresa", "Status", "Prazo Final"]]
        for t in tasks:
            resp_name = t.responsible.name if t.responsible else "N/A"
            resp_comp = t.responsible.company if t.responsible else "N/A"
            p3 = t.date3.strftime('%d/%m/%Y') if t.date3 else "N/A"
            data.append([Paragraph(t.description, styles['Normal']), resp_name, resp_comp, t.status.value, p3])
        tbl = RelTable(data, colWidths=[200, 80, 80, 60, 70])
        tbl.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        elements.append(tbl)
        elements.append(Spacer(1, 24))
        elements.append(Paragraph("Assinaturas:", styles['Heading2']))
        elements.append(Spacer(1, 12))
        participants = set()
        for task in tasks:
            if task.responsible: participants.add(task.responsible)
        for p in participants:
            elements.append(Spacer(1, 10))
            elements.append(Paragraph("__________________________________________", styles['Normal']))
            elements.append(Paragraph(f"{p.name} ({p.company})", styles['Normal']))
        doc.build(elements)
        return filename

# --- UI COMPONENTS ---
class Sidebar(ft.Container):
    def __init__(self, page):
        super().__init__()
        self.m_page = page; self.width = 250; self.bgcolor = ft.Colors.SURFACE_CONTAINER; self.padding = 20
        self.content = ft.Column([
            ft.Container(content=ft.Row([ft.Icon(ft.Icons.POLYMER, color=ft.Colors.CYAN_400, size=30), ft.Text("ATA MASTER", size=20, weight="bold")]), margin=ft.Margin.only(bottom=40)),
            self.nav_item(ft.Icons.DASHBOARD, "Dashboard", "/"),
            self.nav_item(ft.Icons.GROUP, "Grupos & Pessoas", "/management"),
            self.nav_item(ft.Icons.FOLDER, "Minhas Atas", "/meetings"),
            ft.Divider(color=ft.Colors.GREY_800),
            ft.Container(content=ft.Row([ft.Icon(ft.Icons.ADD_CIRCLE, color=ft.Colors.CYAN_400), ft.Text("Nova Reunião", color=ft.Colors.CYAN_400, weight="bold")]), padding=10, border=ft.Border.all(1, ft.Colors.CYAN_900), border_radius=10, on_click=lambda _: self.m_page.go("/new_meeting")),
            ft.Container(expand=True),
            ft.Text("Developed by Daniel Alves Anversi", size=10, color=ft.Colors.GREY_500, italic=True)
        ], expand=True)
    def nav_item(self, icon, text, route): return ft.Container(content=ft.Row([ft.Icon(icon), ft.Text(text)]), padding=10, border_radius=10, on_click=lambda _: self.m_page.go(route), ink=True)

class DashboardView(ft.Column):
    def __init__(self, page):
        super().__init__(expand=True, scroll=ft.ScrollMode.AUTO); self.m_page = page
        try:
            self.controls = [
                ft.Row([
                    ft.Column([
                        ft.Text("Dashboard", size=28, weight="bold"),
                        ft.Text("Visão geral de tarefas e alertas", color=ft.Colors.GREY_400),
                    ]),
                    ft.Row([
                        ft.FilledButton("Exportar Backup", icon=ft.Icons.DOWNLOAD, on_click=self.m_page.on_backup_click),
                        ft.FilledButton("Importar Backup", icon=ft.Icons.UPLOAD, on_click=self.m_page.on_restore_click),
                    ], spacing=10)
                ], alignment="spaceBetween"),
                ft.Divider(height=20, color=ft.Colors.TRANSPARENT),
                self.get_summary_cards(),
                ft.Divider(height=20, color=ft.Colors.TRANSPARENT),
                ft.Text("Alertas Críticos (3º Prazo Vencido)", size=18, weight="bold", color=ft.Colors.RED_400),
                self.get_critical_tasks()
            ]
        except Exception as e: self.controls = [ft.Text(f"Error: {e}")]
    def get_summary_cards(self):
        with get_session() as session:
            total_tasks = session.query(Task).count(); open_tasks = session.query(Task).filter(Task.status == StatusEnum.OPEN).count()
            now = datetime.now(); critical = session.query(Task).filter(Task.status == StatusEnum.OPEN, Task.date3 < now).count()
        return ft.Row([self.summary_card("Total de Tarefas", str(total_tasks), ft.Colors.BLUE_400), self.summary_card("Tarefas Abertas", str(open_tasks), ft.Colors.GREEN_400), self.summary_card("Alertas Críticos", str(critical), ft.Colors.RED_400)], spacing=20)
    def summary_card(self, title, value, color): return ft.Container(content=ft.Column([ft.Text(title, size=14, color=ft.Colors.GREY_400), ft.Text(value, size=30, weight="bold", color=color)]), bgcolor=ft.Colors.SURFACE_CONTAINER, padding=20, border_radius=10, expand=True)
    def get_critical_tasks(self):
        now = datetime.now()
        with get_session() as session:
            tasks = session.query(Task).filter(Task.status == StatusEnum.OPEN, Task.date3 < now).all()
            if not tasks: return ft.Text("Nenhum alerta crítico no momento.", color=ft.Colors.GREY_500)
            task_list = ft.Column(spacing=10)
            for task in tasks:
                def close_task(_, tid=task.id):
                    update_task_status(tid, StatusEnum.CLOSED)
                    self.m_page.go("/"); self.m_page.update() # Refresh
                task_list.controls.append(ft.Container(content=ft.Row([
                    ft.Icon(ft.Icons.WARNING, color=ft.Colors.RED_400),
                    ft.Column([ft.Text(task.description, weight="bold"), ft.Text(f"Responsável: {task.responsible.name if task.responsible else 'N/A'} • Prazo Final: {task.date3.strftime('%d/%m/%Y') if task.date3 else 'N/A'}", size=12)], expand=True),
                    ft.FilledButton("Concluir", icon=ft.Icons.CHECK, on_click=close_task)
                ]), bgcolor=ft.Colors.RED_900, padding=15, border_radius=10))
            return task_list

class ManagementView(ft.Column):
    def __init__(self, page):
        super().__init__(expand=True, scroll=ft.ScrollMode.AUTO); self.m_page = page; self.selected_tab = "groups"; self.refresh(initial=True)
    def refresh(self, initial=False):
        content = ft.Container()
        if self.selected_tab == "groups": content = self.group_tab()
        elif self.selected_tab == "participants": content = self.participant_tab()
        else: content = self.backup_tab()

        self.controls = [
            ft.Text("Gestão de Grupos & Pessoas", size=28, weight="bold"),
            ft.Divider(height=20, color=ft.Colors.TRANSPARENT),
            ft.Row([
                ft.TextButton("Grupos", on_click=lambda _: self.set_tab("groups"), style=ft.ButtonStyle(color=ft.Colors.CYAN_400 if self.selected_tab == "groups" else None)),
                ft.TextButton("Participantes", on_click=lambda _: self.set_tab("participants"), style=ft.ButtonStyle(color=ft.Colors.CYAN_400 if self.selected_tab == "participants" else None)),
                ft.TextButton("Estilos & Backup", on_click=lambda _: self.set_tab("backup"), style=ft.ButtonStyle(color=ft.Colors.CYAN_400 if self.selected_tab == "backup" else None)),
            ]),
            ft.Divider(),
            content
        ]
        if not initial: self.update()
    def set_tab(self, tab): self.selected_tab = tab; self.refresh()
    def group_tab(self):
        groups = get_groups(); group_list = ft.Column(spacing=10)
        for g in groups: group_list.controls.append(ft.ListTile(title=ft.Text(g.name), subtitle=ft.Text(g.description or "Sem descrição"), trailing=ft.IconButton(ft.Icons.DELETE, on_click=lambda _, gid=g.id: self.del_group(gid))))
        name_input = ft.TextField(label="Nome do Grupo", expand=True); desc_input = ft.TextField(label="Descrição", expand=True)
        def add_g(_):
            if name_input.value: create_group(name_input.value, desc_input.value); name_input.value = ""; desc_input.value = ""; self.refresh()
        return ft.Column([ft.Row([name_input, desc_input, ft.FilledButton("Adicionar Grupo", on_click=add_g)]), ft.Divider(), group_list])
    def participant_tab(self):
        participants = get_participants(); p_list = ft.Column(spacing=10)
        for p in participants: p_list.controls.append(ft.ListTile(title=ft.Text(p.name), subtitle=ft.Text(f"{p.company or 'N/A'} • {p.email or 'N/A'}"), trailing=ft.IconButton(ft.Icons.DELETE, on_click=lambda _, pid=p.id: self.del_participant(pid))))
        name_input = ft.TextField(label="Nome", expand=True); email_input = ft.TextField(label="Email", expand=True); company_input = ft.TextField(label="Empresa", expand=True)
        def add_p(_):
            if name_input.value: create_participant(name_input.value, email_input.value, company_input.value); name_input.value = ""; email_input.value = ""; company_input.value = ""; self.refresh()

        excel_actions = ft.Row([
            ft.Text("Importar de Excel (Col A: Nome, Col B: Email):", size=12, color=ft.Colors.GREY_400),
            ft.FilledButton("Selecionar Planilha", icon=ft.Icons.UPLOAD_FILE, on_click=self.m_page.on_excel_click),
            ft.IconButton(ft.Icons.REFRESH, tooltip="Atualizar da última planilha", on_click=self.m_page.on_excel_refresh)
        ], alignment="start")

        return ft.Column([
            ft.Row([name_input, email_input, company_input, ft.FilledButton("Adicionar Participante", on_click=add_p)]),
            ft.Divider(),
            excel_actions,
            ft.Divider(),
            p_list
        ])
    def backup_tab(self):
        def change_theme(mode): self.m_page.theme_mode = ft.ThemeMode.DARK if mode == "DARK" else ft.ThemeMode.LIGHT; self.m_page.update()
        def change_color(color): self.m_page.theme = ft.Theme(color_scheme_seed=color); self.m_page.update()
        return ft.Column([
            ft.Text("Personalização Visual", weight="bold"),
            ft.Row([
                ft.FilledButton("Modo Escuro", icon=ft.Icons.DARK_MODE, on_click=lambda _: change_theme("DARK")),
                ft.FilledButton("Modo Claro", icon=ft.Icons.LIGHT_MODE, on_click=lambda _: change_theme("LIGHT"))
            ]),
            ft.Text("Cor Principal:"),
            ft.Row([
                ft.FilledButton("Azul", bgcolor=ft.Colors.BLUE, color=ft.Colors.WHITE, on_click=lambda _: change_color(ft.Colors.BLUE)),
                ft.FilledButton("Verde", bgcolor=ft.Colors.GREEN, color=ft.Colors.WHITE, on_click=lambda _: change_color(ft.Colors.GREEN)),
                ft.FilledButton("Laranja", bgcolor=ft.Colors.ORANGE, color=ft.Colors.WHITE, on_click=lambda _: change_color(ft.Colors.ORANGE))
            ]),
            ft.Divider(),
            ft.Text("Backup do Banco de Dados", weight="bold"),
            ft.Text("Exporte ou importe todos os dados do aplicativo (atas, grupos, participantes)."),
            ft.Row([
                ft.FilledButton("Exportar Backup", icon=ft.Icons.DOWNLOAD, on_click=self.m_page.on_backup_click),
                ft.FilledButton("Importar Backup", icon=ft.Icons.UPLOAD, on_click=self.m_page.on_restore_click)
            ])
        ], padding=20)
    def del_participant(self, pid): delete_participant(pid); self.refresh()
    def del_group(self, gid): delete_group(gid); self.refresh()

class MeetingsListView(ft.Column):
    def __init__(self, page):
        super().__init__(expand=True, scroll=ft.ScrollMode.AUTO); self.m_page = page; self.refresh(initial=True)
    def refresh(self, initial=False):
        self.controls = [ft.Text("Minhas Atas", size=28, weight="bold"), ft.Text("Atas organizadas por grupo de reunião", color=ft.Colors.GREY_400), ft.Divider(height=20, color=ft.Colors.TRANSPARENT)]
        groups = get_groups()
        if not groups: self.controls.append(ft.Text("Nenhum grupo encontrado. Crie um grupo primeiro.", color=ft.Colors.GREY_500))
        else:
            for g in groups:
                meetings = get_meetings_by_group(g.id)
                group_tile = ft.ExpansionTile(title=ft.Text(g.name, weight="bold", size=18), subtitle=ft.Text(f"{len(meetings)} reuniões", size=12, color=ft.Colors.GREY_400), leading=ft.Icon(ft.Icons.FOLDER_OPEN, color=ft.Colors.CYAN_400), controls=[])
                if not meetings: group_tile.controls.append(ft.ListTile(title=ft.Text("Nenhuma reunião registrada.", size=12, color=ft.Colors.GREY_500)))
                else:
                    for m in meetings: group_tile.controls.append(ft.ListTile(leading=ft.Icon(ft.Icons.EVENT_NOTE, size=20), title=ft.Text(m.title), subtitle=ft.Text(f"{m.date.strftime('%d/%m/%Y')} • {m.location or 'Sem local'}"), trailing=ft.IconButton(ft.Icons.PICTURE_AS_PDF, on_click=lambda _, mid=m.id: self.generate_pdf_click(mid)), on_click=lambda _, mid=m.id: self.view_meeting(mid)))
                self.controls.append(group_tile)
        if not initial: self.update()
    def view_meeting(self, mid): self.m_page.go(f"/meeting/{mid}")
    def generate_pdf_click(self, mid):
        filename = generate_pdf(mid)
        if filename:
            snack = ft.SnackBar(ft.Text(f"PDF gerado: {filename}"))
            self.m_page.overlay.append(snack)
            snack.open = True
            self.m_page.update()

class TaskEditor(ft.Container):
    def __init__(self, on_add_task, page):
        super().__init__()
        self.m_page = page; self.on_add_task = on_add_task; self.participants = get_participants()
        self.desc_input = ft.TextField(label="Descrição da Tarefa", multiline=True, expand=True)
        self.resp_dropdown = ft.Dropdown(label="Responsável", options=[ft.dropdown.Option(str(p.id), p.name) for p in self.participants])
        self.date1 = ft.TextField(label="Prazo 1 (DD/MM/YYYY)", width=150); self.date2 = ft.TextField(label="Prazo 2 (DD/MM/YYYY)", width=150); self.date3 = ft.TextField(label="Prazo 3 (DD/MM/YYYY)", width=150)
        self.content = ft.Column([ft.Text("Adicionar Tarefa", weight="bold"), ft.Row([self.desc_input]), ft.Row([self.resp_dropdown, self.date1, self.date2, self.date3]), ft.FilledButton("Adicionar à Lista", on_click=self.add_clicked)])
    def add_clicked(self, _):
        if self.desc_input.value and self.resp_dropdown.value:
            d1 = self.parse_date(self.date1.value); d2 = self.parse_date(self.date2.value); d3 = self.parse_date(self.date3.value)
            if not d3:
                snack = ft.SnackBar(ft.Text("Erro: Formato de data inválido. Use DD/MM/YYYY."))
                self.m_page.overlay.append(snack); snack.open = True; self.m_page.update()
                return
            self.on_add_task(self.desc_input.value, int(self.resp_dropdown.value), d1, d2, d3)
            self.desc_input.value = ""; self.date1.value = ""; self.date2.value = ""; self.date3.value = ""; self.update()
    def parse_date(self, val):
        try: return datetime.strptime(val, "%d/%m/%Y")
        except: return None

class NewMeetingView(ft.Column):
    def __init__(self, page):
        super().__init__(expand=True, scroll=ft.ScrollMode.AUTO); self.m_page = page; self.tasks_to_add = []; self.groups = get_groups()
        self.group_sel = ft.Dropdown(label="Selecionar Grupo", options=[ft.dropdown.Option(str(g.id), g.name) for g in self.groups], on_select=self.group_changed)
        self.title_input = ft.TextField(label="Título da Reunião", value="Reunião Semanal"); self.location_input = ft.TextField(label="Local", value="Online")
        self.tasks_list_display = ft.Column(); self.task_editor = TaskEditor(self.add_task_to_list, page)
        self.controls = [ft.Text("Nova Reunião", size=28, weight="bold"), ft.Row([self.group_sel, self.title_input, self.location_input]), ft.Divider(), ft.Text("Pauta (Itens Novos e Importados)", size=18, weight="bold"), self.tasks_list_display, ft.Divider(), self.task_editor, ft.Divider(), ft.FilledButton("Salvar e Gerar Ata", icon=ft.Icons.SAVE, on_click=self.save_meeting)]
    def group_changed(self, _):
        if not self.group_sel.value: return
        gid = int(self.group_sel.value); open_tasks = get_open_tasks_for_group(gid); self.tasks_to_add = []
        for t in open_tasks: self.tasks_to_add.append({"id": t.id, "description": t.description, "responsible_id": t.responsible_id, "date1": t.date1, "date2": t.date2, "date3": t.date3, "is_new": False})
        self.refresh_tasks_display()
    def add_task_to_list(self, desc, resp_id, d1, d2, d3):
        lines = desc.split("\n"); adjusted_desc = "\n".join([(l if l.strip().startswith("•") else f"• {l}") for l in lines if l.strip()])
        self.tasks_to_add.append({"description": adjusted_desc, "responsible_id": resp_id, "date1": d1, "date2": d2, "date3": d3, "is_new": True})
        self.refresh_tasks_display()
    def refresh_tasks_display(self):
        self.tasks_list_display.controls.clear()
        for t in self.tasks_to_add:
            p_name = next((p.name for p in self.task_editor.participants if p.id == t["responsible_id"]), "N/A")
            color = ft.Colors.WHITE
            if t.get("date3") and t["date3"] < datetime.now(): color = ft.Colors.RED_400
            def mark_closed(_, tid=t.get("id")):
                if tid: update_task_status(tid, StatusEnum.CLOSED)
                self.tasks_to_add = [task for task in self.tasks_to_add if task.get("id") != tid]
                self.refresh_tasks_display()

            actions = ft.Row([ft.Text("IMPORTADO" if not t["is_new"] else "NOVO", size=10, color=ft.Colors.GREY_500)])
            if not t["is_new"]:
                actions.controls.append(ft.IconButton(ft.Icons.CHECK_CIRCLE, tooltip="Marcar como Concluído", on_click=mark_closed))

            self.tasks_list_display.controls.append(ft.Container(content=ft.Row([ft.Text(t["description"], expand=True, color=color), ft.Text(p_name, width=150), actions]), padding=5))
        self.update()
    def save_meeting(self, _):
        if not self.group_sel.value: return
        gid = int(self.group_sel.value); m = create_meeting(gid, self.title_input.value, location=self.location_input.value); carry_over_tasks(gid, m.id)
        for t in self.tasks_to_add:
            if t["is_new"]: create_task(m.id, t["description"], t["responsible_id"], t["date1"], t["date2"], t["date3"])
        self.m_page.go("/meetings")

class MeetingDetailView(ft.Column):
    def __init__(self, page, mid):
        super().__init__(expand=True, scroll=ft.ScrollMode.AUTO); self.m_page = page; self.mid = int(mid); self.refresh()
    def refresh(self):
        with get_session() as session:
            m = session.query(Meeting).filter(Meeting.id == self.mid).first()
            if not m: self.controls = [ft.Text("Reunião não encontrada.")]; return
            self.controls = [
                ft.Row([ft.IconButton(ft.Icons.ARROW_BACK, on_click=lambda _: self.m_page.go("/meetings")), ft.Text(m.title, size=28, weight="bold")]),
                ft.Text(f"Grupo: {m.group.name} • Data: {m.date.strftime('%d/%m/%Y')}", color=ft.Colors.GREY_400),
                ft.Divider(),
                ft.Text("Tarefas da Reunião", size=18, weight="bold")
            ]
            for t in m.tasks:
                def update_st(tid, status): update_task_status(tid, status); self.refresh(); self.update()
                status_icon = ft.Icons.CHECK_BOX if t.status == StatusEnum.CLOSED else ft.Icons.CHECK_BOX_OUTLINE_BLANK
                self.controls.append(ft.Container(content=ft.Row([
                    ft.IconButton(status_icon, on_click=lambda _, tid=t.id, curr=t.status: update_st(tid, StatusEnum.OPEN if curr == StatusEnum.CLOSED else StatusEnum.CLOSED)),
                    ft.Text(t.description, expand=True, color=ft.Colors.GREY_300 if t.status == StatusEnum.CLOSED else ft.Colors.WHITE),
                    ft.Text(t.responsible.name if t.responsible else "N/A", width=150),
                    ft.Text(t.status.value, size=10, weight="bold", color=ft.Colors.GREEN_400 if t.status == StatusEnum.CLOSED else ft.Colors.BLUE_400)
                ]), padding=5, bgcolor=ft.Colors.SURFACE_CONTAINER if t.status == StatusEnum.OPEN else None))

async def main(page: ft.Page):
    page.title = "AtaMaster Pro"; page.theme_mode = ft.ThemeMode.DARK; page.padding = 0; page.window_min_width = 1100; page.window_min_height = 750
    init_db()

    # Helper for SnackBar
    def show_snack(text):
        snack = ft.SnackBar(ft.Text(text))
        page.overlay.append(snack)
        snack.open = True
        page.update()

    async def import_excel(path):
        if not openpyxl:
            show_snack("Erro: Biblioteca 'openpyxl' não instalada."); return
        try:
            wb = openpyxl.load_workbook(path)
            sheet = wb.active
            count = 0
            for row in sheet.iter_rows(min_row=1, values_only=True):
                name = row[0]
                email = row[1] if len(row) > 1 else None
                if name and str(name).strip():
                    create_participant(str(name), str(email) if email else None)
                    count += 1
            show_snack(f"Importados {count} participantes de {path}")
            # Refresh management view if active
            if page.route == "/management":
                await route_change(None)
        except Exception as ex:
            show_snack(f"Erro ao ler Excel: {ex}")

    page.last_excel_path = None
    page.backup_export_picker = ft.FilePicker()
    page.backup_import_picker = ft.FilePicker()
    page.excel_picker = ft.FilePicker()
    page.overlay.extend([page.backup_export_picker, page.backup_import_picker, page.excel_picker])

    async def on_backup_click(_):
        path = await page.backup_export_picker.save_file(file_name="atamaster_backup.db")
        if path:
            try:
                shutil.copy("atamaster.db", path)
                show_snack(f"Backup exportado para {path}")
            except Exception as ex:
                show_snack(f"Erro ao exportar: {ex}")

    async def on_restore_click(_):
        files = await page.backup_import_picker.pick_files(allowed_extensions=["db"])
        if files:
            try:
                shutil.copy(files[0].path, "atamaster.db")
                show_snack("Backup importado com sucesso! Reinicie o aplicativo.")
            except Exception as ex:
                show_snack(f"Erro ao importar: {ex}")

    async def on_excel_click(_):
        files = await page.excel_picker.pick_files(allowed_extensions=["xlsx", "xls"])
        if files:
            path = files[0].path
            page.last_excel_path = path
            await import_excel(path)

    async def on_excel_refresh(_):
        if page.last_excel_path: await import_excel(page.last_excel_path)
        else: show_snack("Nenhuma planilha selecionada anteriormente.")

    page.on_backup_click = on_backup_click
    page.on_restore_click = on_restore_click
    page.on_excel_click = on_excel_click
    page.on_excel_refresh = on_excel_refresh

    sidebar = Sidebar(page); content_container = ft.Container(expand=True, padding=30, bgcolor=ft.Colors.SURFACE); content_container.content = DashboardView(page)
    async def route_change(route):
        troute = ft.TemplateRoute(page.route)
        if troute.match("/"): content_container.content = DashboardView(page)
        elif troute.match("/management"): content_container.content = ManagementView(page)
        elif troute.match("/meetings"): content_container.content = MeetingsListView(page)
        elif troute.match("/new_meeting"): content_container.content = NewMeetingView(page)
        elif troute.match("/meeting/:id"): content_container.content = MeetingDetailView(page, troute.id)
        page.update()
    page.on_route_change = route_change
    layout = ft.Row([sidebar, ft.VerticalDivider(width=1, color=ft.Colors.GREY_900), ft.Container(content=content_container, expand=True)], expand=True, spacing=0)
    page.add(layout); await page.go(page.route)

if __name__ == "__main__":
    ft.run(main)
