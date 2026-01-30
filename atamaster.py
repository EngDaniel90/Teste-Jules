
import asyncio
import os
import shutil
import flet as ft
from datetime import datetime, date
from sqlalchemy import Column, Integer, String, Boolean, ForeignKey, DateTime, select, update, delete, Table
from sqlalchemy.ext.asyncio import create_async_engine, AsyncSession
from sqlalchemy.orm import sessionmaker, declarative_base, relationship, selectinload
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table as RLTable, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors as rl_colors
import openpyxl

Base = declarative_base()

# Junction Tables
group_participants = Table(
    'group_participants', Base.metadata,
    Column('group_id', Integer, ForeignKey('groups.id'), primary_key=True),
    Column('participant_id', Integer, ForeignKey('participants.id'), primary_key=True)
)

meeting_tasks = Table(
    'meeting_tasks', Base.metadata,
    Column('meeting_id', Integer, ForeignKey('meetings.id'), primary_key=True),
    Column('task_id', Integer, ForeignKey('tasks.id'), primary_key=True)
)

attendance = Table(
    'attendance', Base.metadata,
    Column('meeting_id', Integer, ForeignKey('meetings.id'), primary_key=True),
    Column('participant_id', Integer, ForeignKey('participants.id'), primary_key=True),
    Column('present', Boolean, default=True)
)

class StatusEnum:
    OPEN = "OPEN"
    CLOSED = "CLOSED"

class Group(Base):
    __tablename__ = 'groups'
    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True)
    description = Column(String)
    participants = relationship("Participant", secondary=group_participants, back_populates="groups")
    meetings = relationship("Meeting", back_populates="group")

class Participant(Base):
    __tablename__ = 'participants'
    id = Column(Integer, primary_key=True)
    name = Column(String)
    email = Column(String)
    company = Column(String)
    groups = relationship("Group", secondary=group_participants, back_populates="participants")

class Meeting(Base):
    __tablename__ = 'meetings'
    id = Column(Integer, primary_key=True)
    title = Column(String)
    date = Column(DateTime, default=datetime.now)
    group_id = Column(Integer, ForeignKey('groups.id'))
    group = relationship("Group", back_populates="meetings")
    tasks = relationship("Task", secondary=meeting_tasks, back_populates="tasks")

class Task(Base):
    __tablename__ = 'tasks'
    id = Column(Integer, primary_key=True)
    description = Column(String)
    status = Column(String, default=StatusEnum.OPEN)
    participant_id = Column(Integer, ForeignKey('participants.id'))
    deadline_1 = Column(DateTime, nullable=True)
    deadline_2 = Column(DateTime, nullable=True)
    deadline_3 = Column(DateTime, nullable=True)
    meetings = relationship("Meeting", secondary=meeting_tasks, back_populates="tasks")

class DBManager:
    def __init__(self, db_url="sqlite+aiosqlite:///atamaster.db"):
        self.engine = create_async_engine(db_url)
        self.async_session = sessionmaker(self.engine, expire_on_commit=False, class_=AsyncSession)

    async def init_db(self):
        async with self.engine.begin() as conn:
            await conn.run_sync(Base.metadata.create_all)

    def to_dict(self, obj):
        if obj is None: return None
        return {c.name: getattr(obj, c.name) for c in obj.__table__.columns}

    async def add_group(self, name, description):
        async with self.async_session() as session:
            g = Group(name=name, description=description)
            session.add(g); await session.commit(); return self.to_dict(g)

    async def get_groups(self):
        async with self.async_session() as session:
            res = await session.execute(select(Group))
            return [self.to_dict(g) for g in res.scalars().all()]

    async def add_participant(self, name, email, company):
        async with self.async_session() as session:
            p = Participant(name=name, email=email, company=company)
            session.add(p); await session.commit(); return self.to_dict(p)

    async def get_participants(self):
        async with self.async_session() as session:
            res = await session.execute(select(Participant))
            return [self.to_dict(p) for p in res.scalars().all()]

    async def get_participant(self, p_id):
        async with self.async_session() as session:
            res = await session.execute(select(Participant).filter(Participant.id == p_id))
            return self.to_dict(res.scalars().first())

    async def add_participant_to_group(self, p_id, g_id):
        async with self.async_session() as session:
            stmt = group_participants.insert().values(group_id=g_id, participant_id=p_id)
            await session.execute(stmt); await session.commit()

    async def get_group_participants(self, g_id):
        async with self.async_session() as session:
            res = await session.execute(select(Participant).join(group_participants).filter(group_participants.c.group_id == g_id))
            return [self.to_dict(p) for p in res.scalars().all()]

    async def get_open_tasks_for_group(self, g_id):
        async with self.async_session() as session:
            res = await session.execute(
                select(Task).join(meeting_tasks).join(Meeting).filter(Meeting.group_id == g_id, Task.status == StatusEnum.OPEN).distinct()
            )
            return [self.to_dict(t) for t in res.scalars().all()]

    async def create_meeting(self, title, group_id, task_data, attendance_data):
        async with self.async_session() as session:
            meeting = Meeting(title=title, group_id=group_id)
            session.add(meeting); await session.flush()
            for t_info in task_data:
                if t_info.get("id"):
                    task = await session.get(Task, t_info["id"])
                else:
                    task = Task(
                        description=t_info["description"],
                        participant_id=t_info["participant_id"],
                        deadline_1=t_info.get("deadline_1"),
                        deadline_2=t_info.get("deadline_2"),
                        deadline_3=t_info.get("deadline_3")
                    )
                    session.add(task); await session.flush()
                meeting.tasks.append(task)
            for p_id, present in attendance_data.items():
                await session.execute(attendance.insert().values(meeting_id=meeting.id, participant_id=p_id, present=present))
            await session.commit(); return meeting.id

    async def get_meetings(self):
        async with self.async_session() as session:
            res = await session.execute(select(Meeting).options(selectinload(Meeting.group)).order_by(Meeting.date.desc()))
            return [{**self.to_dict(m), "group_name": m.group.name if m.group else "N/A"} for m in res.scalars().all()]

    async def get_meeting_details(self, m_id):
        async with self.async_session() as session:
            res = await session.execute(select(Meeting).options(selectinload(Meeting.group), selectinload(Meeting.tasks)).filter(Meeting.id == m_id))
            m = res.scalars().first()
            if not m: return None
            att_res = await session.execute(select(attendance).filter(attendance.c.meeting_id == m_id))
            att = {r.participant_id: r.present for r in att_res.all()}
            tasks = [self.to_dict(t) for t in m.tasks]
            return {**self.to_dict(m), "group_name": m.group.name if m.group else "N/A", "tasks": tasks, "attendance": att}

    async def close_task(self, t_id):
        async with self.async_session() as session:
            await session.execute(update(Task).where(Task.id == t_id).values(status=StatusEnum.CLOSED))
            await session.commit()

    async def get_critical_tasks(self):
        async with self.async_session() as session:
            res = await session.execute(select(Task).filter(Task.status == StatusEnum.OPEN, Task.deadline_3 < datetime.now()))
            return [self.to_dict(t) for t in res.scalars().all()]

    async def get_all_open_tasks(self):
        async with self.async_session() as session:
            res = await session.execute(select(Task).filter(Task.status == StatusEnum.OPEN))
            return [self.to_dict(t) for t in res.scalars().all()]

class TaskCard(ft.Container):
    def __init__(self, task, p_name, page):
        super().__init__()
        self.task = task; self.p_name = p_name; self.m_page = page
        self.padding = 15; self.border_radius = 10; self.bgcolor = ft.Colors.SURFACE_VARIANT

        status_color = ft.Colors.GREEN_400 if task['status'] == StatusEnum.CLOSED else ft.Colors.AMBER_400
        is_critical = task['status'] == StatusEnum.OPEN and task['deadline_3'] and task['deadline_3'] < datetime.now()
        if is_critical:
            self.border = ft.Border.all(2, ft.Colors.RED_500)
            status_color = ft.Colors.RED_400

        d1 = task['deadline_1'].strftime('%d/%m') if task['deadline_1'] else "--"
        d2 = task['deadline_2'].strftime('%d/%m') if task['deadline_2'] else "--"
        d3 = task['deadline_3'].strftime('%d/%m') if task['deadline_3'] else "--"

        self.content = ft.Row([
            ft.Column([
                ft.Text(task['description'], weight="bold", size=16),
                ft.Text(f"Responsável: {p_name}", size=12, color=ft.Colors.GREY_400),
                ft.Row([
                    self.date_chip(d1, "P1"), self.date_chip(d2, "P2"), self.date_chip(d3, "P3", critical=is_critical)
                ])
            ], expand=True),
            ft.Container(content=ft.Text(task['status'], size=10, weight="bold", color=ft.Colors.BLACK), bgcolor=status_color, padding=ft.Padding.symmetric(6, 12), border_radius=15)
        ])

    def date_chip(self, text, label, critical=False):
        color = ft.Colors.RED_900 if critical and label == "P3" else ft.Colors.GREY_800
        return ft.Container(content=ft.Text(f"{label}: {text}", size=10), bgcolor=color, padding=ft.Padding.symmetric(2, 6), border_radius=4)

class DashboardView(ft.Column):
    def __init__(self, page):
        super().__init__(expand=True, scroll=ft.ScrollMode.AUTO); self.m_page = page

    async def refresh(self):
        critical = await self.m_page.db.get_critical_tasks()
        open_tasks = await self.m_page.db.get_all_open_tasks()

        self.controls = [
            ft.Container(content=ft.Column([
                ft.Text("Dashboard", size=28, weight="bold"),
                ft.Text("Visão geral de tarefas e alertas", color=ft.Colors.GREY_400),
            ]), padding=ft.Padding.only(bottom=20)),
            ft.Row([
                self.stat_card("Tarefas Críticas", str(len(critical)), ft.Colors.RED_400),
                self.stat_card("Tarefas em Aberto", str(len(open_tasks)), ft.Colors.CYAN_400),
            ], spacing=20),
            ft.Divider(height=40),
            ft.Text("Alertas Críticos (3º Prazo Vencido)", size=20, weight="bold", color=ft.Colors.RED_400),
        ]
        if not critical:
            self.controls.append(ft.Text("Nenhum item crítico no momento.", color=ft.Colors.GREY_500))
        for t in critical:
            p = await self.m_page.db.get_participant(t['participant_id'])
            self.controls.append(TaskCard(t, p['name'] if p else "N/A", self.m_page))
        self.update()

    def stat_card(self, title, value, color):
        return ft.Container(content=ft.Column([ft.Text(title, size=14, color=ft.Colors.GREY_400), ft.Text(value, size=30, weight="bold", color=color)]), bgcolor=ft.Colors.SURFACE_VARIANT, padding=20, border_radius=10, expand=True)

class ManagementView(ft.Column):
    def __init__(self, page):
        super().__init__(expand=True, scroll=ft.ScrollMode.AUTO); self.m_page = page; self.selected_tab = "groups"

    async def refresh(self, initial=False):
        self.controls.clear()
        tabs = ft.Row([
            ft.TextButton("Grupos", on_click=self.goto_groups, style=ft.ButtonStyle(color=ft.Colors.CYAN_400 if self.selected_tab=="groups" else ft.Colors.WHITE)),
            ft.TextButton("Participantes", on_click=self.goto_participants, style=ft.ButtonStyle(color=ft.Colors.CYAN_400 if self.selected_tab=="participants" else ft.Colors.WHITE)),
            ft.TextButton("Backup/Dados", on_click=self.goto_backup, style=ft.ButtonStyle(color=ft.Colors.CYAN_400 if self.selected_tab=="backup" else ft.Colors.WHITE)),
        ])
        content = await self.get_tab_content()
        self.controls = [ft.Text("Gerenciamento", size=28, weight="bold"), tabs, ft.Divider(), content]
        if not initial: self.update()

    async def goto_groups(self, e): await self.set_tab("groups")
    async def goto_participants(self, e): await self.set_tab("participants")
    async def goto_backup(self, e): await self.set_tab("backup")

    async def set_tab(self, tab): self.selected_tab = tab; await self.refresh()

    async def get_tab_content(self):
        if self.selected_tab == "groups": return await self.group_tab()
        if self.selected_tab == "participants": return await self.participant_tab()
        return self.backup_tab()

    async def group_tab(self):
        groups = await self.m_page.db.get_groups()
        g_list = ft.Column([ft.ListTile(title=ft.Text(g['name']), subtitle=ft.Text(g['description'])) for g in groups])
        name_input = ft.TextField(label="Nome do Grupo", expand=True)
        desc_input = ft.TextField(label="Descrição", expand=True)
        async def add_g(e):
            if name_input.value: await self.m_page.db.add_group(name_input.value, desc_input.value); await self.refresh()
        return ft.Column([ft.Row([name_input, desc_input, ft.FilledButton("Adicionar Grupo", on_click=add_g)]), ft.Divider(), g_list])

    async def participant_tab(self):
        ps = await self.m_page.db.get_participants(); gs = await self.m_page.db.get_groups()
        p_list = ft.Column()
        for p in ps:
            p_list.controls.append(ft.ListTile(title=ft.Text(p['name']), subtitle=ft.Text(f"{p['company']} - {p['email']}")))

        name_i = ft.TextField(label="Nome", expand=True); email_i = ft.TextField(label="Email", expand=True); comp_i = ft.TextField(label="Empresa", expand=True)
        group_dropdown = ft.Dropdown(label="Vincular ao Grupo", options=[ft.dropdown.Option(key=str(g['id']), text=g['name']) for g in gs], expand=True)

        async def add_p(e):
            if name_i.value:
                new_p = await self.m_page.db.add_participant(name_i.value, email_i.value, comp_i.value)
                if group_dropdown.value: await self.m_page.db.add_participant_to_group(new_p['id'], int(group_dropdown.value))
                await self.refresh()

        async def import_excel(e):
            res = await self.m_page.excel_picker.pick_files(allowed_extensions=["xlsx"])
            if res and res.files:
                wb = openpyxl.load_workbook(res.files[0].path); ws = wb.active
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row[0]: await self.m_page.db.add_participant(str(row[0]), str(row[1]) if len(row)>1 else "", str(row[2]) if len(row)>2 else "")
                await self.refresh()

        excel_btn = ft.FilledButton("Importar de Excel (A: Nome, B: Email, C: Empresa)", icon=ft.Icons.UPLOAD_FILE, on_click=import_excel)
        return ft.Column([ft.Row([name_i, email_i, comp_i, group_dropdown]), ft.FilledButton("Adicionar Participante", on_click=add_p), ft.Divider(), excel_btn, ft.Divider(), p_list])

    def backup_tab(self):
        return ft.Container(content=ft.Column([
            ft.Text("Manutenção de Dados"),
            ft.Row([
                ft.FilledButton("Exportar Banco (.db)", icon=ft.Icons.SAVE, on_click=self.m_page.run_backup),
                ft.FilledButton("Restaurar Banco (.db)", icon=ft.Icons.RESTORE, on_click=self.m_page.run_restore),
            ])
        ]), padding=20)

class NewMeetingView(ft.Column):
    def __init__(self, page):
        super().__init__(expand=True, scroll=ft.ScrollMode.AUTO); self.m_page = page
        self.temp_tasks = []; self.group_participants = []; self.attendance = {}
        self.deadlines = [None, None, None]

        self.title_i = ft.TextField(label="Título da Reunião", value=f"Reunião {datetime.now().strftime('%d/%m/%Y')}")
        self.group_d = ft.Dropdown(label="Grupo", on_change=self.on_group_select)
        self.attendance_col = ft.Column()

        self.task_desc = ft.TextField(label="Descrição da Tarefa", expand=True)
        self.task_resp = ft.Dropdown(label="Responsável", expand=True)

        self.d1_btn = ft.TextButton("Prazo 1: --", on_click=self.open_dp1)
        self.d2_btn = ft.TextButton("Prazo 2: --", on_click=self.open_dp2)
        self.d3_btn = ft.TextButton("Prazo 3: --", on_click=self.open_dp3)

        self.tasks_list_display = ft.Column()

    async def open_dp1(self, e): self.m_page.active_dp_idx = 0; await self.m_page.date_picker.pick_date()
    async def open_dp2(self, e): self.m_page.active_dp_idx = 1; await self.m_page.date_picker.pick_date()
    async def open_dp3(self, e): self.m_page.active_dp_idx = 2; await self.m_page.date_picker.pick_date()

    async def handle_date_change(self, e):
        idx = self.m_page.active_dp_idx
        d = self.m_page.date_picker.value
        if d:
            self.deadlines[idx] = d
            btns = [self.d1_btn, self.d2_btn, self.d3_btn]
            btns[idx].text = f"Prazo {idx+1}: {d.strftime('%d/%m/%Y')}"
            self.update()

    async def refresh(self):
        gs = await self.m_page.db.get_groups()
        self.group_d.options = [ft.dropdown.Option(key=str(g['id']), text=g['name']) for g in gs]
        self.controls = [
            ft.Text("Nova Reunião", size=28, weight="bold"),
            self.title_i, self.group_d,
            ft.Divider(),
            ft.Text("Chamada / Presença", size=18, weight="bold"),
            self.attendance_col,
            ft.Divider(),
            ft.Text("Novas Tarefas", size=18, weight="bold"),
            ft.Row([self.task_desc, self.task_resp]),
            ft.Row([self.d1_btn, self.d2_btn, self.d3_btn]),
            ft.FilledButton("Adicionar à lista", icon=ft.Icons.ADD, on_click=self.add_task_to_meeting),
            ft.Divider(),
            ft.Text("Itens na Pauta", size=18, weight="bold"),
            self.tasks_list_display,
            ft.Divider(),
            ft.FilledButton("Salvar Ata e Gerar PDF", icon=ft.Icons.SAVE, on_click=self.save_meeting, bgcolor=ft.Colors.GREEN_700, color=ft.Colors.WHITE)
        ]
        self.update()

    async def on_group_select(self, e):
        if not self.group_d.value: return
        g_id = int(self.group_d.value)
        self.group_participants = await self.m_page.db.get_group_participants(g_id)
        self.attendance_col.controls = [ft.Checkbox(label=p['name'], value=True, data=p['id']) for p in self.group_participants]
        self.task_resp.options = [ft.dropdown.Option(key=str(p['id']), text=p['name']) for p in self.group_participants]
        open_tasks = await self.m_page.db.get_open_tasks_for_group(g_id)
        self.temp_tasks = []
        for ot in open_tasks: self.temp_tasks.append({**ot, "from_db": True})
        await self.refresh_tasks()

    async def add_task_to_meeting(self, e):
        if not self.task_desc.value or not self.task_resp.value:
            self.m_page.snack_bar = ft.SnackBar(ft.Text("Preencha descrição e responsável!")); self.m_page.snack_bar.open = True; self.m_page.update()
            return
        self.temp_tasks.append({
            "description": self.task_desc.value, "participant_id": int(self.task_resp.value),
            "deadline_1": self.deadlines[0], "deadline_2": self.deadlines[1], "deadline_3": self.deadlines[2],
            "from_db": False
        })
        self.task_desc.value = ""; self.deadlines = [None, None, None]
        self.d1_btn.text = "Prazo 1: --"; self.d2_btn.text = "Prazo 2: --"; self.d3_btn.text = "Prazo 3: --"
        await self.refresh_tasks()

    async def refresh_tasks(self):
        self.tasks_list_display.controls.clear()
        for i, t in enumerate(self.temp_tasks):
            p = await self.m_page.db.get_participant(t['participant_id'])
            color = ft.Colors.CYAN_200 if t.get("from_db") else ft.Colors.WHITE
            self.tasks_list_display.controls.append(ft.Container(content=ft.Row([
                ft.Text(t["description"], expand=True, color=color),
                ft.Text(p['name'] if p else "N/A", width=150),
                ft.IconButton(ft.Icons.DELETE, on_click=lambda _, idx=i: asyncio.create_task(self.remove_task(idx)))
            ]), padding=5))
        self.update()

    async def remove_task(self, idx): self.temp_tasks.pop(idx); await self.refresh_tasks()

    async def save_meeting(self, e):
        if not self.group_d.value: return
        att_data = {c.data: c.value for c in self.attendance_col.controls}
        m_id = await self.m_page.db.create_meeting(self.title_i.value, int(self.group_d.value), self.temp_tasks, att_data)
        await self.generate_pdf(m_id)
        self.m_page.push_route("/meetings")

    async def generate_pdf(self, m_id):
        m = await self.m_page.db.get_meeting_details(m_id)
        filename = f"Ata_{m_id}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        doc = SimpleDocTemplate(filename, pagesize=A4); styles = getSampleStyleSheet()
        elements = [Paragraph(f"ATA DE REUNIÃO: {m['title']}", styles['Title']), Paragraph(f"Data: {m['date'].strftime('%d/%m/%Y')} | Grupo: {m['group_name']}", styles['Normal']), Spacer(1, 12)]
        elements.append(Paragraph("Participantes / Presença", styles['Heading2']))
        att_rows = [["Nome", "Presença"]]
        for p_id, present in m['attendance'].items():
            p = await self.m_page.db.get_participant(p_id); att_rows.append([p['name'] if p else "N/A", "Sim" if present else "Não"])
        elements.append(RLTable(att_rows, style=TableStyle([('BACKGROUND',(0,0),(-1,0),rl_colors.grey),('TEXTCOLOR',(0,0),(-1,0),rl_colors.whitesmoke)]))); elements.append(Spacer(1, 12))
        elements.append(Paragraph("Tarefas e Acompanhamento", styles['Heading2']))
        task_rows = [["Descrição", "Responsável", "Prazo 3", "Status"]]
        for t in m['tasks']:
            p = await self.m_page.db.get_participant(t['participant_id'])
            d3 = t['deadline_3'].strftime('%d/%m/%Y') if t['deadline_3'] else "--"
            task_rows.append([t['description'], p['name'] if p else "N/A", d3, t['status']])
        elements.append(RLTable(task_rows, style=TableStyle([('GRID', (0,0), (-1,-1), 0.5, rl_colors.black)])))
        elements.append(Spacer(1, 40)); elements.append(Paragraph("Assinaturas:", styles['Heading2']))
        for p_id, present in m['attendance'].items():
            if present:
                p = await self.m_page.db.get_participant(p_id)
                elements.append(Spacer(1, 20)); elements.append(Paragraph("__________________________________________", styles['Normal']))
                elements.append(Paragraph(f"{p['name']} ({p['company']})", styles['Normal']))
        doc.build(elements); os.startfile(filename) if os.name == 'nt' else None

class MeetingsView(ft.Column):
    def __init__(self, page):
        super().__init__(expand=True, scroll=ft.ScrollMode.AUTO); self.m_page = page
    async def refresh(self):
        ms = await self.m_page.db.get_meetings()
        self.controls = [ft.Text("Histórico de Atas", size=28, weight="bold")]
        for m in ms:
            self.controls.append(ft.ListTile(
                title=ft.Text(m['title']), subtitle=ft.Text(f"{m['date'].strftime('%d/%m/%Y')} - Grupo: {m['group_name']}"),
                trailing=ft.Icon(ft.Icons.CHEVRON_RIGHT), on_click=lambda _, mid=m['id']: self.m_page.push_route(f"/meeting/{mid}")
            ))
        self.update()

class MeetingDetailView(ft.Column):
    def __init__(self, page, m_id):
        super().__init__(expand=True, scroll=ft.ScrollMode.AUTO); self.m_page = page; self.m_id = m_id
    async def refresh(self):
        m = await self.m_page.db.get_meeting_details(self.m_id)
        if not m: return
        self.controls = [
            ft.Row([ft.IconButton(ft.Icons.ARROW_BACK, on_click=lambda _: self.m_page.push_route("/meetings")), ft.Text(m['title'], size=28, weight="bold")]),
            ft.Text(f"Data: {m['date'].strftime('%d/%m/%Y')} | Grupo: {m['group_name']}", color=ft.Colors.GREY_400),
            ft.Divider(), ft.Text("Itens desta Reunião", size=20, weight="bold"),
        ]
        for t in m['tasks']:
            p = await self.m_page.db.get_participant(t['participant_id']); actions = ft.Row()
            if t['status'] == StatusEnum.OPEN: actions.controls.append(ft.FilledButton("Fechar Item", on_click=lambda _, tid=t['id']: asyncio.create_task(self.close_item(tid))))
            self.controls.append(ft.Container(content=ft.Row([
                ft.Column([ft.Text(t['description'], weight="bold"), ft.Text(f"Resp: {p['name'] if p else 'N/A'}", size=12)], expand=True),
                ft.Text(t['status'], color=ft.Colors.CYAN_400 if t['status']==StatusEnum.OPEN else ft.Colors.GREEN_400),
                actions
            ]), padding=10, bgcolor=ft.Colors.SURFACE_VARIANT, border_radius=10))
        self.update()
    async def close_item(self, t_id): await self.m_page.db.close_task(t_id); await self.refresh()

class Sidebar(ft.Container):
    def __init__(self, page):
        super().__init__(); self.m_page = page
        self.width = 250; self.bgcolor = ft.Colors.SURFACE_VARIANT; self.padding = 20
        self.theme_btn = ft.IconButton(ft.Icons.LIGHT_MODE if page.theme_mode == ft.ThemeMode.DARK else ft.Icons.DARK_MODE, on_click=self.toggle_theme)
        self.color_dropdown = ft.Dropdown(
            label="Cor Primária", value="cyan", options=[ft.dropdown.Option("cyan", "Cyan"), ft.dropdown.Option("indigo", "Indigo"), ft.dropdown.Option("green", "Green")],
            on_change=self.change_color, text_size=12
        )
        self.content = ft.Column([
            ft.Container(content=ft.Row([ft.Icon(ft.Icons.POLYMER, color=ft.Colors.CYAN_400), ft.Text("ATAMASTER", size=20, weight="bold")]), margin=ft.Margin.only(bottom=40)),
            self.nav_item(ft.Icons.DASHBOARD, "Dashboard", "/"),
            self.nav_item(ft.Icons.FOLDER, "Atas de Reunião", "/meetings"),
            self.nav_item(ft.Icons.GROUP, "Gerenciamento", "/management"),
            ft.Divider(),
            ft.Container(content=ft.Row([ft.Icon(ft.Icons.ADD_CIRCLE, color=ft.Colors.CYAN_400), ft.Text("Nova Reunião", color=ft.Colors.CYAN_400, weight="bold")]), padding=10, border=ft.Border.all(1, ft.Colors.CYAN_900), border_radius=10, on_click=lambda _: self.m_page.push_route("/new_meeting")),
            ft.Divider(),
            ft.Row([ft.Text("Tema", size=12), self.theme_btn], alignment="spaceBetween"),
            self.color_dropdown,
            ft.Container(expand=True),
            ft.Text("Desenvolvido por Daniel Alves Anversi", size=10, italic=True, color=ft.Colors.GREY_500)
        ])
    def nav_item(self, icon, text, route): return ft.Container(content=ft.Row([ft.Icon(icon), ft.Text(text)]), padding=10, border_radius=10, on_click=lambda _, r=route: self.m_page.push_route(r), ink=True)
    def toggle_theme(self, e):
        self.m_page.theme_mode = ft.ThemeMode.LIGHT if self.m_page.theme_mode == ft.ThemeMode.DARK else ft.ThemeMode.DARK
        self.theme_btn.icon = ft.Icons.LIGHT_MODE if self.m_page.theme_mode == ft.ThemeMode.DARK else ft.Icons.DARK_MODE
        self.m_page.update()
    def change_color(self, e):
        c = self.color_dropdown.value
        self.m_page.theme = ft.Theme(color_scheme_seed=c); self.m_page.update()

async def main(page: ft.Page):
    page.title = "AtaMaster Pro"; page.theme_mode = ft.ThemeMode.DARK; page.padding = 0; page.theme = ft.Theme(color_scheme_seed="cyan")
    page.db = DBManager(); await page.db.init_db()
    page.excel_picker = ft.FilePicker(); page.backup_picker = ft.FilePicker()
    page.date_picker = ft.DatePicker(on_change=lambda e: asyncio.create_task(page.current_view.handle_date_change(e)) if hasattr(page.current_view, 'handle_date_change') else None)
    page.overlay.extend([page.excel_picker, page.backup_picker, page.date_picker])

    async def run_backup(e=None):
        path = await page.backup_picker.save_file(file_name="atamaster_backup.db")
        if path: shutil.copy("atamaster.db", path)
    page.run_backup = run_backup
    async def run_restore(e=None):
        res = await page.backup_picker.pick_files(allowed_extensions=["db"])
        if res and res.files: shutil.copy(res.files[0].path, "atamaster.db"); await page.db.init_db()
    page.run_restore = run_restore

    sidebar = Sidebar(page); content_container = ft.Container(expand=True, padding=30, bgcolor=ft.Colors.SURFACE)
    async def route_change(e):
        content_container.content = None; page.update(); route = page.route
        if route == "/": view = DashboardView(page)
        elif route == "/meetings": view = MeetingsView(page)
        elif route == "/management": view = ManagementView(page)
        elif route == "/new_meeting": view = NewMeetingView(page)
        elif route.startswith("/meeting/"): view = MeetingDetailView(page, int(route.split("/")[-1]))
        else: view = ft.Text("404 Not Found")
        page.current_view = view; content_container.content = view
        if hasattr(view, 'refresh'): await view.refresh()
        page.update()
    page.on_route_change = route_change
    layout = ft.Row([sidebar, ft.VerticalDivider(width=1, color=ft.Colors.GREY_900), content_container], expand=True, spacing=0)
    page.add(layout); page.push_route("/")

if __name__ == "__main__":
    ft.run(main)
