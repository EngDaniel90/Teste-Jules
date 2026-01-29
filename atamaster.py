import os
import shutil
import enum
from datetime import datetime
from sqlalchemy import create_engine, Column, Integer, String, Text, DateTime, ForeignKey, Enum, Table
from sqlalchemy.orm import declarative_base, sessionmaker, relationship
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table as RelTable, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
import flet as ft

# ==========================================
# 1. BANCO DE DADOS E MODELOS
# ==========================================

Base = declarative_base()

class StatusEnum(enum.Enum):
    OPEN = "ABERTO"
    IN_PROGRESS = "EM ANDAMENTO"
    CLOSED = "CONCLUÍDO"

# Tabela de associação Reunião <-> Tarefa
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
    role = Column(String(100)) # Cargo
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
    meeting_origin_id = Column(Integer, ForeignKey('meetings.id')) # Onde nasceu
    description = Column(Text, nullable=False)
    responsible_id = Column(Integer, ForeignKey('participants.id'))
    status = Column(Enum(StatusEnum), default=StatusEnum.OPEN)
    date_deadline = Column(DateTime) # Simplificado para um prazo principal
    meetings = relationship("Meeting", secondary=meeting_tasks, back_populates="tasks")
    responsible = relationship("Participant", back_populates="tasks")

# Configuração SQLite
DB_URL = "sqlite:///atamaster_pro.db"
engine = create_engine(DB_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

def init_db():
    Base.metadata.create_all(bind=engine)

def get_session():
    return SessionLocal()

# --- CRUD Operations ---
def db_create_group(name, description):
    with get_session() as s:
        g = Group(name=name, description=description)
        s.add(g); s.commit(); s.refresh(g)
        return g

def db_get_groups():
    with get_session() as s: return s.query(Group).all()

def db_delete_group(id):
    with get_session() as s:
        s.query(Group).filter(Group.id==id).delete(); s.commit()

def db_create_participant(name, email, company, role):
    with get_session() as s:
        p = Participant(name=name, email=email, company=company, role=role)
        s.add(p); s.commit(); s.refresh(p)
        return p

def db_get_participants():
    with get_session() as s: return s.query(Participant).all()

def db_delete_participant(id):
    with get_session() as s:
        s.query(Participant).filter(Participant.id==id).delete(); s.commit()

def db_create_meeting(group_id, title, date_obj, location):
    with get_session() as s:
        m = Meeting(group_id=group_id, title=title, date=date_obj, location=location)
        s.add(m); s.commit(); s.refresh(m)
        return m.id

def db_get_meetings_summary():
    with get_session() as s:
        return s.query(Meeting).order_by(Meeting.date.desc()).all()

def db_get_meeting_details(meeting_id):
    with get_session() as s:
        m = s.query(Meeting).filter(Meeting.id == meeting_id).first()
        if not m: return None, []
        # Force load tasks
        tasks = [t for t in m.tasks]
        return m, tasks

def db_add_task(meeting_id, description, resp_id, deadline):
    with get_session() as s:
        t = Task(meeting_origin_id=meeting_id, description=description, 
                 responsible_id=resp_id, date_deadline=deadline)
        s.add(t)
        # Link to meeting
        m = s.query(Meeting).filter(Meeting.id == meeting_id).first()
        m.tasks.append(t)
        s.commit()

def db_link_existing_task(meeting_id, task_id):
    with get_session() as s:
        m = s.query(Meeting).filter(Meeting.id == meeting_id).first()
        t = s.query(Task).filter(Task.id == task_id).first()
        if m and t and t not in m.tasks:
            m.tasks.append(t)
            s.commit()

def db_get_open_tasks(group_id):
    # Retorna tarefas abertas de reuniões anteriores deste grupo
    with get_session() as s:
        return s.query(Task).join(meeting_tasks).join(Meeting)\
            .filter(Meeting.group_id == group_id)\
            .filter(Task.status != StatusEnum.CLOSED).distinct().all()

def db_update_task_status(task_id, new_status_str):
    with get_session() as s:
        t = s.query(Task).filter(Task.id == task_id).first()
        if t:
            if new_status_str == "ABERTO": t.status = StatusEnum.OPEN
            elif new_status_str == "EM ANDAMENTO": t.status = StatusEnum.IN_PROGRESS
            elif new_status_str == "CONCLUÍDO": t.status = StatusEnum.CLOSED
            s.commit()

# ==========================================
# 2. GERAÇÃO DE PDF
# ==========================================
def generate_pdf_report(meeting_id):
    with get_session() as session:
        meeting = session.query(Meeting).filter(Meeting.id == meeting_id).first()
        if not meeting: return None
        tasks = meeting.tasks
        
        filename = f"ATA_{meeting.id}_{datetime.now().strftime('%Y%m%d')}.pdf"
        doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('TitleCustom', parent=styles['Title'], fontSize=24, spaceAfter=20, textColor=colors.HexColor("#1A237E"))
        sub_style = ParagraphStyle('SubCustom', parent=styles['Normal'], fontSize=12, leading=16)
        
        elements = []
        
        # Cabeçalho
        elements.append(Paragraph(f"ATA DE REUNIÃO", title_style))
        elements.append(Paragraph(f"<b>Grupo:</b> {meeting.group.name}", sub_style))
        elements.append(Paragraph(f"<b>Assunto:</b> {meeting.title}", sub_style))
        elements.append(Paragraph(f"<b>Data:</b> {meeting.date.strftime('%d/%m/%Y')} | <b>Local:</b> {meeting.location}", sub_style))
        elements.append(Spacer(1, 20))
        
        # Tabela de Ações
        if tasks:
            elements.append(Paragraph("Plano de Ação", styles['Heading2']))
            data = [["O QUE (Descrição)", "QUEM (Responsável)", "STATUS", "PRAZO"]]
            
            for t in tasks:
                resp = t.responsible.name if t.responsible else "N/A"
                status = t.status.value
                prazo = t.date_deadline.strftime('%d/%m/%Y') if t.date_deadline else "-"
                
                # Wrap text
                p_desc = Paragraph(t.description, styles['Normal'])
                data.append([p_desc, resp, status, prazo])
            
            t = RelTable(data, colWidths=[240, 100, 90, 80])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A237E")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                ('TOPPADDING', (0, 0), (-1, 0), 10),
            ]))
            elements.append(t)
        else:
            elements.append(Paragraph("Nenhuma tarefa ou deliberação registrada nesta reunião.", styles['Normal']))

        elements.append(Spacer(1, 40))
        
        # Assinaturas
        elements.append(Paragraph("Participantes Presentes", styles['Heading3']))
        participants = set([t.responsible for t in tasks if t.responsible])
        
        sig_data = []
        row = []
        for i, p in enumerate(participants):
            row.append(f"___________________________\n{p.name}\n{p.company or ''}")
            if len(row) == 2:
                sig_data.append(row)
                row = []
        if row: sig_data.append(row)
        
        if sig_data:
            sig_table = RelTable(sig_data, colWidths=[250, 250])
            sig_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 30),
            ]))
            elements.append(sig_table)
            
        doc.build(elements)
        return filename

# ==========================================
# 3. COMPONENTES UI
# ==========================================

class AppTheme:
    primary = ft.Colors.INDIGO
    secondary = ft.Colors.CYAN_ACCENT
    bg = ft.Colors.GREY_900
    surface = ft.Colors.GREY_800
    text_muted = ft.Colors.GREY_400

class StatCard(ft.Container):
    def __init__(self, title, value, icon, color):
        super().__init__()
        self.expand = True
        self.padding = 20
        self.bgcolor = AppTheme.surface
        self.border_radius = 12
        self.shadow = ft.BoxShadow(spread_radius=1, blur_radius=10, color=ft.Colors.BLACK12)
        self.content = ft.Row([
            ft.Container(
                content=ft.Icon(icon, size=30, color=color),
                bgcolor=ft.Colors.with_opacity(0.1, color),
                padding=10, border_radius=10
            ),
            ft.Column([
                ft.Text(title, size=12, color=AppTheme.text_muted, weight="bold"),
                ft.Text(str(value), size=24, weight="bold")
            ], spacing=2)
        ], alignment=ft.MainAxisAlignment.START)

class Sidebar(ft.Container):
    def __init__(self, page, nav_callback):
        super().__init__()
        self.page_ref = page
        self.nav_callback = nav_callback
        self.width = 250
        self.bgcolor = "#111418"  # Darker than surface
        self.padding = ft.padding.symmetric(vertical=30, horizontal=15)
        
        self.menu_items = [
            {"icon": ft.Icons.DASHBOARD_ROUNDED, "label": "Dashboard", "idx": 0},
            {"icon": ft.Icons.FOLDER_SHARED_ROUNDED, "label": "Gestão (Grupos/Pessoas)", "idx": 1},
            {"icon": ft.Icons.EVENT_NOTE_ROUNDED, "label": "Atas & Reuniões", "idx": 2},
        ]
        
        self.nav_controls = []
        for item in self.menu_items:
            self.nav_controls.append(self._build_nav_item(item))

        self.content = ft.Column([
            ft.Row([
                ft.Icon(ft.Icons.POLYMER_SHARP, color=ft.Colors.INDIGO_400, size=34),
                ft.Text("ATA MASTER", size=22, weight="bold", font_family="Roboto")
            ], alignment="center", spacing=10),
            ft.Divider(color=ft.Colors.GREY_800, height=40),
            ft.Column(self.nav_controls, spacing=5),
            ft.Spacer(),
            ft.Container(
                content=ft.Row([
                    ft.Icon(ft.Icons.SETTINGS, size=16, color=ft.Colors.GREY_500),
                    ft.Text("Versão 2.0 Pro", size=12, color=ft.Colors.GREY_500)
                ], alignment="center"),
                padding=10
            )
        ])

    def _build_nav_item(self, item):
        return ft.Container(
            content=ft.Row([
                ft.Icon(item['icon'], size=20, color=ft.Colors.WHITE70),
                ft.Text(item['label'], size=14, weight="w500", color=ft.Colors.WHITE)
            ], spacing=15),
            padding=12,
            border_radius=8,
            on_click=lambda e: self.nav_callback(item['idx']),
            ink=True,
            data=item['idx']
        )
    
    def set_active(self, idx):
        for control in self.nav_controls:
            is_active = control.data == idx
            control.bgcolor = ft.Colors.INDIGO_900 if is_active else None
            control.content.controls[0].color = ft.Colors.CYAN_200 if is_active else ft.Colors.WHITE70
            control.content.controls[1].color = ft.Colors.CYAN_100 if is_active else ft.Colors.WHITE
        self.update()

# ==========================================
# 4. VIEW - DASHBOARD
# ==========================================
class DashboardView(ft.Column):
    def __init__(self, page):
        super().__init__(expand=True, scroll=ft.ScrollMode.HIDDEN)
        self.page_ref = page
        self.refresh_data()

    def refresh_data(self):
        with get_session() as s:
            total_meetings = s.query(Meeting).count()
            open_tasks = s.query(Task).filter(Task.status == StatusEnum.OPEN).count()
            participants = s.query(Participant).count()
            
            # Tarefas recentes
            recent_tasks = s.query(Task).order_by(Task.id.desc()).limit(5).all()

        self.controls = [
            ft.Text("Visão Geral", size=28, weight="bold"),
            ft.Divider(color="transparent", height=10),
            ft.Row([
                StatCard("Reuniões Realizadas", total_meetings, ft.Icons.MEETING_ROOM, ft.Colors.BLUE),
                StatCard("Ações em Aberto", open_tasks, ft.Icons.ASSIGNMENT_LATE, ft.Colors.ORANGE),
                StatCard("Participantes", participants, ft.Icons.PEOPLE, ft.Colors.GREEN),
            ], spacing=20),
            ft.Divider(color="transparent", height=30),
            ft.Text("Últimas Ações Cadastradas", size=20, weight="bold"),
            ft.Container(
                content=self._build_task_table(recent_tasks),
                bgcolor=AppTheme.surface,
                border_radius=10,
                padding=10
            )
        ]
        self.update()

    def _build_task_table(self, tasks):
        if not tasks: return ft.Text("Nenhuma tarefa encontrada.", italic=True, color=ft.Colors.GREY_500)
        
        rows = []
        for t in tasks:
            status_color = ft.Colors.RED if t.status == StatusEnum.OPEN else ft.Colors.GREEN
            rows.append(ft.DataRow(cells=[
                ft.DataCell(ft.Text(str(t.id))),
                ft.DataCell(ft.Text(t.description, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS, width=300)),
                ft.DataCell(ft.Container(
                    content=ft.Text(t.status.value, size=10, weight="bold"),
                    bgcolor=ft.Colors.with_opacity(0.1, status_color),
                    padding=5, border_radius=5
                )),
                ft.DataCell(ft.Text(t.responsible.name if t.responsible else "-")),
            ]))
        
        return ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("ID")),
                ft.DataColumn(ft.Text("Descrição")),
                ft.DataColumn(ft.Text("Status")),
                ft.DataColumn(ft.Text("Responsável")),
            ],
            rows=rows,
            border=ft.border.all(1, ft.Colors.GREY_800),
            vertical_lines=ft.border.BorderSide(1, ft.Colors.GREY_800),
            horizontal_lines=ft.border.BorderSide(1, ft.Colors.GREY_800),
        )

# ==========================================
# 5. VIEW - GESTÃO (Grupos & Pessoas)
# ==========================================
class ManagementView(ft.Column):
    def __init__(self, page):
        super().__init__(expand=True, scroll=ft.ScrollMode.AUTO)
        self.page_ref = page
        
        # Elementos UI para Grupos
        self.group_table = ft.DataTable(columns=[
            ft.DataColumn(ft.Text("ID")),
            ft.DataColumn(ft.Text("Nome")),
            ft.DataColumn(ft.Text("Descrição")),
            ft.DataColumn(ft.Text("Ações"))
        ], width=float('inf'))
        
        # Elementos UI para Pessoas
        self.people_table = ft.DataTable(columns=[
            ft.DataColumn(ft.Text("Nome")),
            ft.DataColumn(ft.Text("Empresa")),
            ft.DataColumn(ft.Text("Email")),
            ft.DataColumn(ft.Text("Ações"))
        ], width=float('inf'))

        # Tabs
        self.tabs = ft.Tabs(
            selected_index=0,
            animation_duration=300,
            tabs=[
                ft.Tab(text="Grupos", icon=ft.Icons.GROUP_WORK, content=ft.Column([
                    ft.Row([ft.FilledButton("Novo Grupo", icon=ft.Icons.ADD, on_click=self.open_group_dialog)], alignment="end"),
                    ft.Container(content=self.group_table, bgcolor=AppTheme.surface, border_radius=10, padding=10)
                ], spacing=20, scroll=True)),
                
                ft.Tab(text="Participantes", icon=ft.Icons.PERSON, content=ft.Column([
                    ft.Row([ft.FilledButton("Novo Participante", icon=ft.Icons.ADD, on_click=self.open_person_dialog)], alignment="end"),
                    ft.Container(content=self.people_table, bgcolor=AppTheme.surface, border_radius=10, padding=10)
                ], spacing=20, scroll=True))
            ],
            expand=True
        )
        
        self.controls = [
            ft.Text("Gestão de Cadastros", size=28, weight="bold"),
            ft.Divider(height=20, color="transparent"),
            self.tabs
        ]
        
        self.load_data()

    def load_data(self):
        # Load Groups
        self.group_table.rows.clear()
        for g in db_get_groups():
            self.group_table.rows.append(ft.DataRow(cells=[
                ft.DataCell(ft.Text(str(g.id))),
                ft.DataCell(ft.Text(g.name, weight="bold")),
                ft.DataCell(ft.Text(g.description or "-")),
                ft.DataCell(ft.IconButton(ft.Icons.DELETE, icon_color=ft.Colors.RED_400, on_click=lambda e, gid=g.id: self.delete_group(gid)))
            ]))
            
        # Load People
        self.people_table.rows.clear()
        for p in db_get_participants():
            self.people_table.rows.append(ft.DataRow(cells=[
                ft.DataCell(ft.Text(p.name, weight="bold")),
                ft.DataCell(ft.Text(p.company or "-")),
                ft.DataCell(ft.Text(p.email or "-")),
                ft.DataCell(ft.IconButton(ft.Icons.DELETE, icon_color=ft.Colors.RED_400, on_click=lambda e, pid=p.id: self.delete_person(pid)))
            ]))
        self.update()

    def delete_group(self, id):
        db_delete_group(id); self.load_data()
    
    def delete_person(self, id):
        db_delete_participant(id); self.load_data()

    # --- Dialogs ---
    def open_group_dialog(self, e):
        nm = ft.TextField(label="Nome do Grupo"); ds = ft.TextField(label="Descrição")
        def save(e):
            if nm.value:
                db_create_group(nm.value, ds.value)
                dlg.open = False; self.page_ref.update(); self.load_data()
        
        dlg = ft.AlertDialog(title=ft.Text("Novo Grupo"), content=ft.Column([nm, ds], height=150),
            actions=[ft.TextButton("Cancelar", on_click=lambda e: self.page_ref.close(dlg)), ft.FilledButton("Salvar", on_click=save)])
        self.page_ref.open(dlg)

    def open_person_dialog(self, e):
        nm = ft.TextField(label="Nome"); em = ft.TextField(label="Email")
        cp = ft.TextField(label="Empresa"); rl = ft.TextField(label="Cargo")
        def save(e):
            if nm.value:
                db_create_participant(nm.value, em.value, cp.value, rl.value)
                dlg.open = False; self.page_ref.update(); self.load_data()
        
        dlg = ft.AlertDialog(title=ft.Text("Novo Participante"), content=ft.Column([nm, em, cp, rl], height=250),
            actions=[ft.TextButton("Cancelar", on_click=lambda e: self.page_ref.close(dlg)), ft.FilledButton("Salvar", on_click=save)])
        self.page_ref.open(dlg)

# ==========================================
# 6. VIEW - ATAS & REUNIÕES (Complexo)
# ==========================================
class MeetingsView(ft.Column):
    def __init__(self, page):
        super().__init__(expand=True, scroll=ft.ScrollMode.AUTO)
        self.page_ref = page
        self.current_view = "list" # list | form | details
        self.render_list()

    def render_list(self):
        meetings = db_get_meetings_summary()
        
        rows = []
        for m in meetings:
            rows.append(ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(m.date.strftime("%d/%m/%Y"), weight="bold")),
                    ft.DataCell(ft.Text(m.group.name if m.group else "N/A")),
                    ft.DataCell(ft.Text(m.title)),
                    ft.DataCell(ft.Text(f"{len(m.tasks)} ações")),
                    ft.DataCell(ft.Row([
                        ft.IconButton(ft.Icons.VISIBILITY, tooltip="Ver Detalhes", on_click=lambda e, mid=m.id: self.render_details(mid)),
                        ft.IconButton(ft.Icons.PICTURE_AS_PDF, icon_color=ft.Colors.RED_400, tooltip="Gerar PDF", on_click=lambda e, mid=m.id: self.trigger_pdf(mid))
                    ]))
                ],
                on_select_changed=lambda e, mid=m.id: self.render_details(mid)
            ))
        
        table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Data")),
                ft.DataColumn(ft.Text("Grupo")),
                ft.DataColumn(ft.Text("Assunto")),
                ft.DataColumn(ft.Text("Qtd. Ações")),
                ft.DataColumn(ft.Text("Opções")),
            ],
            rows=rows,
            width=float('inf'),
            heading_row_color=ft.Colors.BLACK12,
        )

        self.controls = [
            ft.Row([
                ft.Text("Minhas Atas", size=28, weight="bold"),
                ft.FilledButton("Nova Reunião", icon=ft.Icons.ADD, on_click=self.render_form)
            ], alignment="spaceBetween"),
            ft.Divider(height=20, color="transparent"),
            ft.Container(content=table, bgcolor=AppTheme.surface, border_radius=10, padding=10)
        ]
        self.update()

    def trigger_pdf(self, mid):
        f = generate_pdf_report(mid)
        if f:
            self.page_ref.show_snack_bar(ft.SnackBar(ft.Text(f"PDF Gerado: {f}", color="white"), bgcolor="green"))
        else:
            self.page_ref.show_snack_bar(ft.SnackBar(ft.Text("Erro ao gerar PDF"), bgcolor="red"))

    # --- CRIAÇÃO DE REUNIÃO ---
    def render_form(self, e=None):
        groups = db_get_groups()
        if not groups:
            self.page_ref.show_snack_bar(ft.SnackBar(ft.Text("Crie um Grupo primeiro!")))
            return

        # Inputs
        dd_group = ft.Dropdown(label="Grupo", options=[ft.dropdown.Option(g.id, g.name) for g in groups], expand=True)
        txt_title = ft.TextField(label="Assunto da Reunião", expand=True)
        txt_date = ft.TextField(label="Data (DD/MM/YYYY)", value=datetime.now().strftime("%d/%m/%Y"), width=150)
        txt_loc = ft.TextField(label="Local", value="Online / Teams", expand=True)
        
        # Área de tarefas pendentes (Carry over)
        lv_pending = ft.ListView(height=150, spacing=5)
        selected_pending_ids = []

        def on_group_change(e):
            if not dd_group.value: return
            lv_pending.controls.clear()
            pending = db_get_open_tasks(int(dd_group.value))
            if not pending:
                lv_pending.controls.append(ft.Text("Nenhuma pendência anterior.", italic=True))
            else:
                for t in pending:
                    chk = ft.Checkbox(label=f"{t.description} ({t.responsible.name if t.responsible else '?'})", value=True)
                    chk.data = t.id
                    lv_pending.controls.append(chk)
            lv_pending.update()

        dd_group.on_change = on_group_change

        def save_meeting(e):
            if not dd_group.value or not txt_title.value: return
            
            try:
                dt_obj = datetime.strptime(txt_date.value, "%d/%m/%Y")
            except:
                self.page_ref.show_snack_bar(ft.SnackBar(ft.Text("Data inválida"))); return

            # Criar Reunião
            mid = db_create_meeting(int(dd_group.value), txt_title.value, dt_obj, txt_loc.value)
            
            # Associar tarefas antigas selecionadas
            for c in lv_pending.controls:
                if isinstance(c, ft.Checkbox) and c.value:
                    db_link_existing_task(mid, c.data)
            
            self.render_details(mid) # Ir para detalhes para adicionar novas tarefas

        self.controls = [
            ft.Row([
                ft.IconButton(ft.Icons.ARROW_BACK, on_click=lambda e: self.render_list()),
                ft.Text("Nova Ata de Reunião", size=24, weight="bold")
            ]),
            ft.Divider(),
            ft.Container(
                content=ft.Column([
                    ft.Text("1. Dados Básicos", weight="bold"),
                    ft.Row([dd_group, txt_date]),
                    ft.Row([txt_title, txt_loc]),
                    ft.Divider(),
                    ft.Text("2. Revisão de Pendências (Atas Anteriores)", weight="bold"),
                    ft.Container(content=lv_pending, bgcolor=ft.Colors.BLACK26, border_radius=5, padding=10),
                    ft.Divider(),
                    ft.Row([ft.FilledButton("Criar e Ir para Ações", icon=ft.Icons.CHECK, on_click=save_meeting)], alignment="end")
                ], spacing=15),
                bgcolor=AppTheme.surface, padding=20, border_radius=10
            )
        ]
        self.update()

    # --- DETALHES & ADICIONAR AÇÕES ---
    def render_details(self, mid):
        meeting, tasks = db_get_meeting_details(mid)
        if not meeting: return

        # Tabela de tarefas da reunião
        def update_status(e, tid):
            db_update_task_status(tid, e.control.value)
            self.page_ref.show_snack_bar(ft.SnackBar(ft.Text("Status atualizado!")))

        task_rows = []
        for t in tasks:
            dd_status = ft.Dropdown(
                value=t.status.value,
                options=[ft.dropdown.Option(x.value) for x in StatusEnum],
                text_size=12, height=40, content_padding=5,
                on_change=lambda e, tid=t.id: update_status(e, tid)
            )
            
            task_rows.append(ft.DataRow(cells=[
                ft.DataCell(ft.Text(t.description, width=250, max_lines=2)),
                ft.DataCell(ft.Text(t.responsible.name if t.responsible else "-")),
                ft.DataCell(ft.Text(t.date_deadline.strftime("%d/%m") if t.date_deadline else "-")),
                ft.DataCell(dd_status),
            ]))

        task_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("O Que")),
                ft.DataColumn(ft.Text("Quem")),
                ft.DataColumn(ft.Text("Quando")),
                ft.DataColumn(ft.Text("Status")),
            ],
            rows=task_rows,
            border=ft.border.all(1, ft.Colors.GREY_800)
        )

        def add_task_dialog(e):
            parts = db_get_participants()
            txt_desc = ft.TextField(label="Descrição", multiline=True)
            dd_resp = ft.Dropdown(label="Responsável", options=[ft.dropdown.Option(p.id, p.name) for p in parts])
            txt_dline = ft.TextField(label="Prazo (DD/MM/YYYY)", value=datetime.now().strftime("%d/%m/%Y"))
            
            def save_task(e):
                if not txt_desc.value: return
                try: d_obj = datetime.strptime(txt_dline.value, "%d/%m/%Y")
                except: d_obj = None
                
                db_add_task(meeting.id, txt_desc.value, int(dd_resp.value) if dd_resp.value else None, d_obj)
                dlg.open = False
                self.page_ref.update()
                self.render_details(meeting.id)

            dlg = ft.AlertDialog(title=ft.Text("Nova Ação / Deliberação"), content=ft.Column([txt_desc, dd_resp, txt_dline], height=250),
                actions=[ft.TextButton("Cancelar", on_click=lambda e: self.page_ref.close(dlg)), ft.FilledButton("Adicionar", on_click=save_task)])
            self.page_ref.open(dlg)

        self.controls = [
            ft.Row([
                ft.IconButton(ft.Icons.ARROW_BACK, on_click=lambda e: self.render_list()),
                ft.Column([
                    ft.Text(meeting.title, size=20, weight="bold"),
                    ft.Text(f"{meeting.group.name} - {meeting.date.strftime('%d/%m/%Y')}", color="grey")
                ])
            ]),
            ft.Divider(),
            ft.Row([
                ft.Text("Plano de Ação", size=18, weight="bold"),
                ft.FilledButton("Adicionar Item", icon=ft.Icons.ADD_TASK, on_click=add_task_dialog)
            ], alignment="spaceBetween"),
            ft.Container(content=task_table, bgcolor=AppTheme.surface, border_radius=10, padding=10, margin=ft.margin.only(top=10)),
            ft.Divider(height=30, color="transparent"),
            ft.FilledButton("Exportar PDF", icon=ft.Icons.PICTURE_AS_PDF, style=ft.ButtonStyle(bgcolor=ft.Colors.RED_700), on_click=lambda e: self.trigger_pdf(meeting.id))
        ]
        self.update()


# ==========================================
# 7. APLICAÇÃO PRINCIPAL
# ==========================================

def main(page: ft.Page):
    page.title = "AtaMaster Pro Enterprise"
    page.theme_mode = ft.ThemeMode.DARK
    page.theme = ft.Theme(color_scheme_seed=ft.Colors.INDIGO)
    page.padding = 0
    page.window_min_width = 1000
    page.window_min_height = 700

    # Inicializa DB
    init_db()

    # File Picker (Backup) - Sem tipagem estrita para evitar erros de versão
    def on_export_result(e):
        if e.path:
            try:
                shutil.copy(DB_URL.replace("sqlite:///", ""), e.path)
                page.show_snack_bar(ft.SnackBar(ft.Text("Backup realizado com sucesso!")))
            except Exception as ex:
                page.show_snack_bar(ft.SnackBar(ft.Text(f"Erro: {ex}")))

    def on_import_result(e):
        if e.files:
            try:
                shutil.copy(e.files[0].path, DB_URL.replace("sqlite:///", ""))
                page.show_snack_bar(ft.SnackBar(ft.Text("Backup restaurado! Reinicie o app.")))
            except Exception as ex:
                page.show_snack_bar(ft.SnackBar(ft.Text(f"Erro: {ex}")))

    export_picker = ft.FilePicker(on_result=on_export_result)
    import_picker = ft.FilePicker(on_result=on_import_result)
    page.overlay.extend([export_picker, import_picker])

    # Área de Conteúdo
    content_area = ft.Container(expand=True, padding=30)
    
    def navigate(idx):
        content_area.content = None
        if idx == 0: content_area.content = DashboardView(page)
        elif idx == 1: content_area.content = ManagementView(page)
        elif idx == 2: content_area.content = MeetingsView(page)
        
        sidebar.set_active(idx)
        content_area.update()

    sidebar = Sidebar(page, navigate)

    # Layout Principal
    page.add(
        ft.Row(
            [
                sidebar,
                ft.VerticalDivider(width=1, color=ft.Colors.GREY_900),
                content_area
            ],
            expand=True,
            spacing=0
        )
    )

    # Iniciar no Dashboard
    navigate(0)

if __name__ == "__main__":
    ft.app(target=main)
