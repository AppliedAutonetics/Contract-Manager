import os
import re
import json
import difflib
from datetime import datetime, date
from io import BytesIO

from flask import (Flask, render_template, redirect, url_for, flash, request,
                   jsonify, send_file, abort, session)
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.utils import secure_filename

from models import db, User, Client, ContractTemplate, Contract, ContractRevision, ContractFieldValue, AuditLog

# PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# Word doc handling
try:
    from docx import Document as DocxDocument
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

app = Flask(__name__)

# Configuration
app.config['SECRET_KEY'] = os.environ.get('SESSION_SECRET', 'dev-secret-key-change-in-prod')
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'postgresql://localhost/contracts')
if app.config['SQLALCHEMY_DATABASE_URI'].startswith('postgres://'):
    app.config['SQLALCHEMY_DATABASE_URI'] = app.config['SQLALCHEMY_DATABASE_URI'].replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload
ALLOWED_EXTENSIONS = {'txt', 'pdf', 'docx', 'doc'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Initialize extensions
db.init_app(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Please log in to access this page.'
login_manager.login_message_category = 'info'


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


# ─── Helpers ──────────────────────────────────────────────────────────────────

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_template_fields(content):
    """Find all {{FIELD_NAME}} markers in template content."""
    pattern = r'\{\{([A-Z_][A-Z0-9_]*)\}\}'
    fields = list(dict.fromkeys(re.findall(pattern, content)))
    return fields


def apply_field_values(content, field_values):
    """Replace {{FIELD_NAME}} markers with actual values."""
    for name, value in field_values.items():
        content = content.replace('{{' + name + '}}', value or f'[{name}]')
    return content


def generate_contract_number():
    year = datetime.utcnow().year
    count = Contract.query.filter(
        db.extract('year', Contract.created_at) == year
    ).count() + 1
    return f'CTR-{year}-{count:04d}'


def log_action(action, resource_type=None, resource_id=None, contract_id=None, details=None):
    log = AuditLog(
        user_id=current_user.id if current_user.is_authenticated else None,
        action=action,
        resource_type=resource_type,
        resource_id=resource_id,
        contract_id=contract_id,
        details=details,
        ip_address=request.remote_addr,
        created_at=datetime.utcnow()
    )
    db.session.add(log)


def extract_text_from_file(file_path, filename):
    """Extract text content from uploaded file."""
    ext = filename.rsplit('.', 1)[1].lower() if '.' in filename else ''
    try:
        if ext == 'txt':
            with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                return f.read()
        elif ext == 'docx' and DOCX_AVAILABLE:
            doc = DocxDocument(file_path)
            return '\n'.join([para.text for para in doc.paragraphs])
        else:
            return f'[File: {filename} - content extraction not supported for this format. File stored at {file_path}]'
    except Exception as e:
        return f'[Error reading file: {str(e)}]'


# ─── PDF Generation ───────────────────────────────────────────────────────────

def generate_pdf(contract, revision=None):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=inch,
        leftMargin=inch,
        topMargin=inch,
        bottomMargin=inch
    )

    styles = getSampleStyleSheet()
    story = []

    # Custom styles
    title_style = ParagraphStyle(
        'ContractTitle',
        parent=styles['Title'],
        fontSize=20,
        spaceAfter=6,
        textColor=colors.HexColor('#1a2742'),
        alignment=TA_CENTER
    )
    subtitle_style = ParagraphStyle(
        'Subtitle',
        parent=styles['Normal'],
        fontSize=11,
        spaceAfter=4,
        textColor=colors.HexColor('#64748b'),
        alignment=TA_CENTER
    )
    heading_style = ParagraphStyle(
        'SectionHeading',
        parent=styles['Heading2'],
        fontSize=13,
        spaceBefore=16,
        spaceAfter=8,
        textColor=colors.HexColor('#1a2742'),
        borderPad=4
    )
    body_style = ParagraphStyle(
        'ContractBody',
        parent=styles['Normal'],
        fontSize=10,
        leading=16,
        spaceAfter=8,
        textColor=colors.HexColor('#374151')
    )
    info_style = ParagraphStyle(
        'InfoStyle',
        parent=styles['Normal'],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor('#374151')
    )

    # Header
    story.append(Spacer(1, 0.2 * inch))
    story.append(Paragraph(contract.title, title_style))
    story.append(Paragraph(f'Contract #{contract.contract_number}', subtitle_style))
    story.append(Spacer(1, 0.1 * inch))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#1a2742')))
    story.append(Spacer(1, 0.2 * inch))

    # Contract details table
    status_map = {'draft': 'Draft', 'in_review': 'In Review', 'approved': 'Approved', 'finalized': 'Finalized', 'expired': 'Expired'}
    details_data = [
        ['Client:', contract.client.name, 'Status:', status_map.get(contract.status, contract.status)],
        ['Created:', contract.created_at.strftime('%B %d, %Y'), 'Creator:', contract.creator.full_name],
    ]
    if contract.start_date:
        details_data.append(['Start Date:', contract.start_date.strftime('%B %d, %Y'), 'End Date:', contract.end_date.strftime('%B %d, %Y') if contract.end_date else 'N/A'])
    if contract.value:
        details_data.append(['Contract Value:', f'${float(contract.value):,.2f}', '', ''])

    details_table = Table(details_data, colWidths=[1.3 * inch, 2.5 * inch, 1.3 * inch, 2.5 * inch])
    details_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (2, 0), (2, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#374151')),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#1a2742')),
        ('TEXTCOLOR', (2, 0), (2, -1), colors.HexColor('#1a2742')),
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [colors.HexColor('#f8fafc'), colors.white]),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
        ('PADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(details_table)
    story.append(Spacer(1, 0.3 * inch))

    # Contract content
    content = ''
    if revision:
        content = revision.content
        story.append(Paragraph(f'Version {revision.version_number}' + (' (FINAL)' if revision.is_finalized else ''), subtitle_style))
    elif contract.latest_revision:
        content = contract.latest_revision.content

    if content:
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#e2e8f0')))
        story.append(Spacer(1, 0.15 * inch))

        for para_text in content.split('\n\n'):
            para_text = para_text.strip()
            if not para_text:
                story.append(Spacer(1, 0.1 * inch))
                continue

            # Detect section headings (ALL CAPS or ends with colon)
            lines = para_text.split('\n')
            first_line = lines[0].strip()
            if first_line.isupper() and len(first_line) < 80:
                story.append(Paragraph(first_line, heading_style))
                if len(lines) > 1:
                    rest = ' '.join(lines[1:]).strip()
                    if rest:
                        story.append(Paragraph(rest, body_style))
            else:
                # Normal paragraph - replace newlines with spaces
                text = para_text.replace('\n', ' ')
                story.append(Paragraph(text, body_style))

    # Footer info
    story.append(Spacer(1, 0.4 * inch))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#e2e8f0')))
    story.append(Spacer(1, 0.1 * inch))
    gen_time = datetime.utcnow().strftime('%B %d, %Y at %H:%M UTC')
    story.append(Paragraph(f'Generated on {gen_time}', subtitle_style))
    if contract.finalized_at:
        story.append(Paragraph(f'Finalized on {contract.finalized_at.strftime("%B %d, %Y")}', subtitle_style))

    doc.build(story)
    buffer.seek(0)
    return buffer


# ─── Auth Routes ──────────────────────────────────────────────────────────────

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        remember = request.form.get('remember') == 'on'

        user = User.query.filter_by(email=email).first()
        if user and user.check_password(password):
            user.last_login = datetime.utcnow()
            db.session.commit()
            login_user(user, remember=remember)
            next_page = request.args.get('next')
            flash(f'Welcome back, {user.full_name}!', 'success')
            return redirect(next_page or url_for('dashboard'))
        else:
            flash('Invalid email or password.', 'error')

    return render_template('auth/login.html')


@app.route('/register', methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        full_name = request.form.get('full_name', '').strip()
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        confirm = request.form.get('confirm_password', '')

        if not all([full_name, email, password]):
            flash('All fields are required.', 'error')
        elif password != confirm:
            flash('Passwords do not match.', 'error')
        elif len(password) < 8:
            flash('Password must be at least 8 characters.', 'error')
        elif User.query.filter_by(email=email).first():
            flash('An account with this email already exists.', 'error')
        else:
            user = User(full_name=full_name, email=email)
            user.set_password(password)
            db.session.add(user)
            db.session.commit()
            login_user(user)
            flash(f'Account created! Welcome, {full_name}.', 'success')
            return redirect(url_for('dashboard'))

    return render_template('auth/register.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('You have been logged out.', 'info')
    return redirect(url_for('login'))


# ─── Dashboard ────────────────────────────────────────────────────────────────

@app.route('/')
@login_required
def dashboard():
    total_contracts = Contract.query.filter_by(created_by=current_user.id).count()
    by_status = {
        'draft': Contract.query.filter_by(created_by=current_user.id, status='draft').count(),
        'in_review': Contract.query.filter_by(created_by=current_user.id, status='in_review').count(),
        'approved': Contract.query.filter_by(created_by=current_user.id, status='approved').count(),
        'finalized': Contract.query.filter_by(created_by=current_user.id, status='finalized').count(),
    }
    recent_contracts = Contract.query.filter_by(created_by=current_user.id).order_by(Contract.updated_at.desc()).limit(8).all()
    total_clients = Client.query.filter_by(created_by=current_user.id).count()
    total_templates = ContractTemplate.query.filter_by(is_active=True, created_by=current_user.id).count()
    recent_activity = AuditLog.query.filter_by(user_id=current_user.id).order_by(AuditLog.created_at.desc()).limit(10).all()

    return render_template('dashboard.html',
        total_contracts=total_contracts,
        by_status=by_status,
        recent_contracts=recent_contracts,
        total_clients=total_clients,
        total_templates=total_templates,
        recent_activity=recent_activity
    )


# ─── Client Routes ────────────────────────────────────────────────────────────

@app.route('/clients')
@login_required
def clients_list():
    q = request.args.get('q', '')
    ctype = request.args.get('type', '')
    query = Client.query.filter_by(created_by=current_user.id)
    if q:
        query = query.filter(
            db.or_(Client.name.ilike(f'%{q}%'), Client.company.ilike(f'%{q}%'), Client.email.ilike(f'%{q}%'))
        )
    if ctype:
        query = query.filter_by(client_type=ctype)
    clients = query.order_by(Client.name).all()
    return render_template('clients/list.html', clients=clients, q=q, ctype=ctype)


@app.route('/clients/new', methods=['GET', 'POST'])
@login_required
def clients_new():
    if request.method == 'POST':
        client = Client(
            name=request.form.get('name', '').strip(),
            company=request.form.get('company', '').strip(),
            email=request.form.get('email', '').strip(),
            phone=request.form.get('phone', '').strip(),
            address=request.form.get('address', '').strip(),
            notes=request.form.get('notes', '').strip(),
            client_type=request.form.get('client_type', 'client'),
            created_by=current_user.id
        )
        if not client.name:
            flash('Name is required.', 'error')
            return render_template('clients/new.html', client=client)

        db.session.add(client)
        log_action('create_client', 'client', None, details=f'Created client: {client.name}')
        db.session.commit()
        flash(f'Client "{client.name}" created successfully.', 'success')
        return redirect(url_for('clients_detail', client_id=client.id))

    return render_template('clients/new.html', client=None)


@app.route('/clients/<int:client_id>')
@login_required
def clients_detail(client_id):
    client = Client.query.get_or_404(client_id)
    if client.created_by != current_user.id:
        abort(403)
    contracts = client.contracts.order_by(Contract.updated_at.desc()).all()
    return render_template('clients/detail.html', client=client, contracts=contracts)


@app.route('/clients/<int:client_id>/edit', methods=['GET', 'POST'])
@login_required
def clients_edit(client_id):
    client = Client.query.get_or_404(client_id)
    if client.created_by != current_user.id:
        abort(403)
    if request.method == 'POST':
        client.name = request.form.get('name', '').strip()
        client.company = request.form.get('company', '').strip()
        client.email = request.form.get('email', '').strip()
        client.phone = request.form.get('phone', '').strip()
        client.address = request.form.get('address', '').strip()
        client.notes = request.form.get('notes', '').strip()
        client.client_type = request.form.get('client_type', 'client')
        client.updated_at = datetime.utcnow()
        log_action('update_client', 'client', client.id, details=f'Updated client: {client.name}')
        db.session.commit()
        flash('Client updated successfully.', 'success')
        return redirect(url_for('clients_detail', client_id=client.id))
    return render_template('clients/new.html', client=client)


# ─── Template Routes ──────────────────────────────────────────────────────────

@app.route('/templates')
@login_required
def templates_list():
    q = request.args.get('q', '')
    ttype = request.args.get('type', '')
    query = ContractTemplate.query.filter_by(is_active=True, created_by=current_user.id)
    if q:
        query = query.filter(ContractTemplate.name.ilike(f'%{q}%'))
    if ttype:
        query = query.filter_by(template_type=ttype)
    templates = query.order_by(ContractTemplate.updated_at.desc()).all()
    return render_template('templates/list.html', templates=templates, q=q, ttype=ttype)


@app.route('/templates/new', methods=['GET', 'POST'])
@login_required
def templates_new():
    if request.method == 'POST':
        content = request.form.get('content', '').strip()

        # Handle file upload
        if 'file' in request.files and request.files['file'].filename:
            file = request.files['file']
            if allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], f'tpl_{datetime.utcnow().strftime("%Y%m%d%H%M%S")}_{filename}')
                file.save(file_path)
                content = extract_text_from_file(file_path, filename)

        if not content:
            flash('Template content is required. Enter text or upload a file.', 'error')
            return render_template('templates/new.html', template=None)

        fields = extract_template_fields(content)
        template = ContractTemplate(
            name=request.form.get('name', '').strip(),
            description=request.form.get('description', '').strip(),
            template_type=request.form.get('template_type', 'contract'),
            content=content,
            created_by=current_user.id
        )
        template.fields = fields

        if not template.name:
            flash('Template name is required.', 'error')
            return render_template('templates/new.html', template=template)

        db.session.add(template)
        log_action('create_template', 'template', None, details=f'Created template: {template.name}')
        db.session.commit()
        flash(f'Template "{template.name}" created with {len(fields)} fields detected.', 'success')
        return redirect(url_for('templates_detail', template_id=template.id))

    return render_template('templates/new.html', template=None)


@app.route('/templates/<int:template_id>')
@login_required
def templates_detail(template_id):
    template = ContractTemplate.query.get_or_404(template_id)
    if template.created_by != current_user.id:
        abort(403)
    return render_template('templates/detail.html', template=template)


@app.route('/templates/<int:template_id>/edit', methods=['GET', 'POST'])
@login_required
def templates_edit(template_id):
    template = ContractTemplate.query.get_or_404(template_id)
    if template.created_by != current_user.id:
        abort(403)
    if request.method == 'POST':
        content = request.form.get('content', '').strip()

        if 'file' in request.files and request.files['file'].filename:
            file = request.files['file']
            if allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], f'tpl_{datetime.utcnow().strftime("%Y%m%d%H%M%S")}_{filename}')
                file.save(file_path)
                content = extract_text_from_file(file_path, filename)

        template.name = request.form.get('name', '').strip()
        template.description = request.form.get('description', '').strip()
        template.template_type = request.form.get('template_type', 'contract')
        template.content = content
        template.fields = extract_template_fields(content)
        template.updated_at = datetime.utcnow()
        log_action('update_template', 'template', template.id, details=f'Updated template: {template.name}')
        db.session.commit()
        flash('Template updated.', 'success')
        return redirect(url_for('templates_detail', template_id=template.id))

    return render_template('templates/new.html', template=template)


@app.route('/templates/<int:template_id>/delete', methods=['POST'])
@login_required
def templates_delete(template_id):
    template = ContractTemplate.query.get_or_404(template_id)
    if template.created_by != current_user.id:
        abort(403)
    template.is_active = False
    log_action('delete_template', 'template', template.id, details=f'Deleted template: {template.name}')
    db.session.commit()
    flash('Template deleted.', 'success')
    return redirect(url_for('templates_list'))


@app.route('/api/templates/<int:template_id>/fields')
@login_required
def template_fields_api(template_id):
    template = ContractTemplate.query.get_or_404(template_id)
    if template.created_by != current_user.id:
        abort(403)
    return jsonify({'fields': template.fields, 'content': template.content})


# ─── Contract Routes ──────────────────────────────────────────────────────────

@app.route('/contracts')
@login_required
def contracts_list():
    q = request.args.get('q', '')
    status = request.args.get('status', '')
    client_id = request.args.get('client_id', '')
    query = Contract.query.filter_by(created_by=current_user.id)
    if q:
        query = query.filter(
            db.or_(Contract.title.ilike(f'%{q}%'), Contract.contract_number.ilike(f'%{q}%'))
        )
    if status:
        query = query.filter_by(status=status)
    if client_id:
        query = query.filter_by(client_id=int(client_id))
    contracts = query.order_by(Contract.updated_at.desc()).all()
    clients = Client.query.filter_by(created_by=current_user.id).order_by(Client.name).all()
    return render_template('contracts/list.html', contracts=contracts, clients=clients,
                           q=q, status=status, client_id=client_id)


@app.route('/contracts/new', methods=['GET', 'POST'])
@login_required
def contracts_new():
    clients = Client.query.filter_by(created_by=current_user.id).order_by(Client.name).all()
    templates = ContractTemplate.query.filter_by(is_active=True, created_by=current_user.id).order_by(ContractTemplate.name).all()

    if request.method == 'POST':
        template_id = request.form.get('template_id')
        template_id = int(template_id) if template_id else None

        # Validate that the submitted client and template belong to the current user
        submitted_client_id = request.form.get('client_id')
        if not submitted_client_id:
            flash('Client is required.', 'error')
            return render_template('contracts/new.html', clients=clients, templates=templates, contract=None)
        submitted_client = Client.query.filter_by(id=int(submitted_client_id), created_by=current_user.id).first()
        if not submitted_client:
            abort(403)

        if template_id:
            submitted_template = ContractTemplate.query.filter_by(id=template_id, created_by=current_user.id, is_active=True).first()
            if not submitted_template:
                abort(403)

        start_date = None
        end_date = None
        try:
            sd = request.form.get('start_date', '')
            ed = request.form.get('end_date', '')
            if sd:
                start_date = date.fromisoformat(sd)
            if ed:
                end_date = date.fromisoformat(ed)
        except ValueError:
            pass

        value = None
        try:
            v = request.form.get('value', '').strip()
            if v:
                value = float(v.replace(',', ''))
        except ValueError:
            pass

        contract = Contract(
            title=request.form.get('title', '').strip(),
            client_id=submitted_client.id,
            template_id=template_id,
            status='draft',
            notes=request.form.get('notes', '').strip(),
            start_date=start_date,
            end_date=end_date,
            value=value,
            created_by=current_user.id,
            contract_number=generate_contract_number()
        )

        if not contract.title or not contract.client_id:
            flash('Title and client are required.', 'error')
            return render_template('contracts/new.html', clients=clients, templates=templates, contract=contract)

        db.session.add(contract)
        db.session.flush()

        # Create initial revision from template or manual content
        initial_content = request.form.get('initial_content', '').strip()
        if template_id and not initial_content:
            template = ContractTemplate.query.get(template_id)
            if template:
                initial_content = template.content

                # Save field values
                for field_name in template.fields:
                    field_val = request.form.get(f'field_{field_name}', '').strip()
                    if field_val:
                        fv = ContractFieldValue(
                            contract_id=contract.id,
                            field_name=field_name,
                            field_value=field_val
                        )
                        db.session.add(fv)

                # Apply fields to content
                field_values = {f: request.form.get(f'field_{f}', '') for f in template.fields}
                initial_content = apply_field_values(initial_content, field_values)

        if initial_content:
            revision = ContractRevision(
                contract_id=contract.id,
                version_number=1,
                content=initial_content,
                changes_summary='Initial version',
                created_by=current_user.id
            )
            db.session.add(revision)

        log_action('create_contract', 'contract', contract.id, contract_id=contract.id,
                   details=f'Created contract: {contract.title}')
        db.session.commit()
        flash(f'Contract "{contract.title}" created.', 'success')
        return redirect(url_for('contracts_detail', contract_id=contract.id))

    return render_template('contracts/new.html', clients=clients, templates=templates, contract=None)


@app.route('/contracts/<int:contract_id>')
@login_required
def contracts_detail(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    if contract.created_by != current_user.id:
        abort(403)
    revisions = contract.revisions.order_by(ContractRevision.version_number.desc()).all()
    field_values = {fv.field_name: fv.field_value for fv in contract.field_values}
    audit = contract.audit_logs.order_by(AuditLog.created_at.desc()).limit(20).all()
    return render_template('contracts/detail.html', contract=contract, revisions=revisions,
                           field_values=field_values, audit=audit)


@app.route('/contracts/<int:contract_id>/edit', methods=['GET', 'POST'])
@login_required
def contracts_edit(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    if contract.created_by != current_user.id:
        abort(403)
    clients = Client.query.filter_by(created_by=current_user.id).order_by(Client.name).all()

    if request.method == 'POST':
        # Validate that the submitted client belongs to the current user
        submitted_client_id = request.form.get('client_id')
        if not submitted_client_id:
            flash('Client is required.', 'error')
            return render_template('contracts/edit.html', contract=contract, clients=clients)
        submitted_client = Client.query.filter_by(id=int(submitted_client_id), created_by=current_user.id).first()
        if not submitted_client:
            abort(403)

        contract.title = request.form.get('title', '').strip()
        contract.client_id = submitted_client.id
        contract.status = request.form.get('status', contract.status)
        contract.notes = request.form.get('notes', '').strip()
        contract.updated_at = datetime.utcnow()

        try:
            sd = request.form.get('start_date', '')
            ed = request.form.get('end_date', '')
            contract.start_date = date.fromisoformat(sd) if sd else None
            contract.end_date = date.fromisoformat(ed) if ed else None
        except ValueError:
            pass

        try:
            v = request.form.get('value', '').strip()
            contract.value = float(v.replace(',', '')) if v else None
        except ValueError:
            pass

        if contract.status == 'finalized' and not contract.finalized_at:
            contract.finalized_at = datetime.utcnow()

        log_action('update_contract', 'contract', contract.id, contract_id=contract.id,
                   details=f'Updated contract: {contract.title}, status={contract.status}')
        db.session.commit()
        flash('Contract updated.', 'success')
        return redirect(url_for('contracts_detail', contract_id=contract.id))

    return render_template('contracts/edit.html', contract=contract, clients=clients)


@app.route('/contracts/<int:contract_id>/revision/new', methods=['GET', 'POST'])
@login_required
def revisions_new(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    if contract.created_by != current_user.id:
        abort(403)

    if request.method == 'POST':
        content = request.form.get('content', '').strip()
        changes_summary = request.form.get('changes_summary', '').strip()
        file_name = None

        # Handle file upload
        if 'file' in request.files and request.files['file'].filename:
            file = request.files['file']
            if allowed_file(file.filename):
                file_name = secure_filename(file.filename)
                ts = datetime.utcnow().strftime('%Y%m%d%H%M%S')
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], f'rev_{contract_id}_{ts}_{file_name}')
                file.save(file_path)
                content = extract_text_from_file(file_path, file_name)
            else:
                flash('Invalid file type. Allowed: txt, pdf, docx', 'error')
                return redirect(request.url)

        if not content:
            flash('Revision content is required.', 'error')
            return redirect(request.url)

        last_rev = contract.revisions.order_by(ContractRevision.version_number.desc()).first()
        next_version = (last_rev.version_number + 1) if last_rev else 1

        revision = ContractRevision(
            contract_id=contract.id,
            version_number=next_version,
            content=content,
            changes_summary=changes_summary or f'Version {next_version}',
            file_name=file_name,
            created_by=current_user.id
        )
        db.session.add(revision)
        contract.updated_at = datetime.utcnow()

        log_action('add_revision', 'contract_revision', None, contract_id=contract.id,
                   details=f'Added revision v{next_version} to contract {contract.title}')
        db.session.commit()
        flash(f'Revision v{next_version} added.', 'success')
        return redirect(url_for('contracts_detail', contract_id=contract.id))

    latest = contract.latest_revision
    return render_template('contracts/revision_new.html', contract=contract, latest=latest)


@app.route('/contracts/<int:contract_id>/compare')
@login_required
def contracts_compare(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    if contract.created_by != current_user.id:
        abort(403)
    rev1_id = request.args.get('rev1', type=int)
    rev2_id = request.args.get('rev2', type=int)

    revisions = contract.revisions.order_by(ContractRevision.version_number).all()
    if len(revisions) < 2:
        flash('You need at least 2 revisions to compare.', 'warning')
        return redirect(url_for('contracts_detail', contract_id=contract_id))

    rev1 = ContractRevision.query.get_or_404(rev1_id) if rev1_id else revisions[-2]
    if rev1_id and rev1.contract_id != contract_id:
        abort(403)
    rev2 = ContractRevision.query.get_or_404(rev2_id) if rev2_id else revisions[-1]
    if rev2_id and rev2.contract_id != contract_id:
        abort(403)

    # Generate HTML diff
    differ = difflib.HtmlDiff(wrapcolumn=80)
    diff_html = differ.make_table(
        rev1.content.splitlines(keepends=True),
        rev2.content.splitlines(keepends=True),
        fromdesc=f'Version {rev1.version_number}',
        todesc=f'Version {rev2.version_number}',
        context=True,
        numlines=3
    )

    # Also generate unified diff for stats
    unified = list(difflib.unified_diff(
        rev1.content.splitlines(),
        rev2.content.splitlines()
    ))
    added = sum(1 for l in unified if l.startswith('+') and not l.startswith('+++'))
    removed = sum(1 for l in unified if l.startswith('-') and not l.startswith('---'))

    return render_template('contracts/compare.html', contract=contract, rev1=rev1, rev2=rev2,
                           revisions=revisions, diff_html=diff_html, added=added, removed=removed)


@app.route('/contracts/<int:contract_id>/finalize', methods=['POST'])
@login_required
def contracts_finalize(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    if contract.created_by != current_user.id:
        abort(403)
    revision_id = request.form.get('revision_id', type=int)

    if revision_id:
        revision = ContractRevision.query.get_or_404(revision_id)
        if revision.contract_id != contract_id:
            abort(403)
    else:
        revision = contract.latest_revision

    if not revision:
        flash('No revision found to finalize.', 'error')
        return redirect(url_for('contracts_detail', contract_id=contract_id))

    # Mark revision as finalized
    revision.is_finalized = True
    contract.status = 'finalized'
    contract.finalized_at = datetime.utcnow()
    contract.updated_at = datetime.utcnow()

    log_action('finalize_contract', 'contract', contract.id, contract_id=contract.id,
               details=f'Finalized contract: {contract.title}, revision v{revision.version_number}')
    db.session.commit()
    flash(f'Contract "{contract.title}" has been finalized (v{revision.version_number}).', 'success')
    return redirect(url_for('contracts_detail', contract_id=contract_id))


@app.route('/contracts/<int:contract_id>/pdf')
@login_required
def contracts_pdf(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    if contract.created_by != current_user.id:
        abort(403)
    revision_id = request.args.get('revision_id', type=int)
    if revision_id:
        revision = ContractRevision.query.get_or_404(revision_id)
        if revision.contract_id != contract_id:
            abort(403)
    else:
        revision = contract.latest_revision

    pdf_buffer = generate_pdf(contract, revision)
    filename = f'{contract.contract_number}_{contract.title.replace(" ", "_")}'
    if revision:
        filename += f'_v{revision.version_number}'
    filename += '.pdf'

    log_action('export_pdf', 'contract', contract.id, contract_id=contract.id,
               details=f'Exported PDF for contract: {contract.title}')
    db.session.commit()

    return send_file(pdf_buffer, as_attachment=True, download_name=filename,
                     mimetype='application/pdf')


@app.route('/contracts/<int:contract_id>/status', methods=['POST'])
@login_required
def contracts_status(contract_id):
    contract = Contract.query.get_or_404(contract_id)
    if contract.created_by != current_user.id:
        abort(403)
    new_status = request.form.get('status')
    valid_statuses = ['draft', 'in_review', 'approved', 'finalized', 'expired']
    if new_status in valid_statuses:
        old_status = contract.status
        contract.status = new_status
        contract.updated_at = datetime.utcnow()
        if new_status == 'finalized' and not contract.finalized_at:
            contract.finalized_at = datetime.utcnow()
        log_action('status_change', 'contract', contract.id, contract_id=contract.id,
                   details=f'Status changed: {old_status} → {new_status}')
        db.session.commit()
        flash(f'Status updated to {new_status.replace("_", " ").title()}.', 'success')
    return redirect(url_for('contracts_detail', contract_id=contract_id))


# ─── Template Parsing API ─────────────────────────────────────────────────────

@app.route('/api/detect-fields', methods=['POST'])
@login_required
def detect_fields_api():
    content = request.json.get('content', '')
    fields = extract_template_fields(content)
    return jsonify({'fields': fields})


# ─── Init ─────────────────────────────────────────────────────────────────────

def create_tables():
    with app.app_context():
        db.create_all()


if __name__ == '__main__':
    create_tables()
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port, debug=False)
