from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
import json

db = SQLAlchemy()


class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(255), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(50), default='user')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_login = db.Column(db.DateTime)

    contracts = db.relationship('Contract', foreign_keys='Contract.created_by', backref='creator', lazy='dynamic')
    audit_logs = db.relationship('AuditLog', backref='user', lazy='dynamic')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return f'<User {self.email}>'


class Client(db.Model):
    __tablename__ = 'clients'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False)
    company = db.Column(db.String(255))
    email = db.Column(db.String(255))
    phone = db.Column(db.String(100))
    address = db.Column(db.Text)
    notes = db.Column(db.Text)
    client_type = db.Column(db.String(50), default='client')  # client or vendor
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    contracts = db.relationship('Contract', backref='client', lazy='dynamic')

    def __repr__(self):
        return f'<Client {self.name}>'


class ContractTemplate(db.Model):
    __tablename__ = 'contract_templates'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False)
    description = db.Column(db.Text)
    template_type = db.Column(db.String(100), default='contract')  # contract or sow
    content = db.Column(db.Text, nullable=False)
    fields_json = db.Column(db.Text, default='[]')  # JSON list of field names found in template
    is_active = db.Column(db.Boolean, default=True)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    contracts = db.relationship('Contract', backref='template', lazy='dynamic')
    creator = db.relationship('User', foreign_keys=[created_by])

    @property
    def fields(self):
        return json.loads(self.fields_json or '[]')

    @fields.setter
    def fields(self, value):
        self.fields_json = json.dumps(value)

    def __repr__(self):
        return f'<ContractTemplate {self.name}>'


class Contract(db.Model):
    __tablename__ = 'contracts'
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(255), nullable=False)
    contract_number = db.Column(db.String(100), unique=True)
    client_id = db.Column(db.Integer, db.ForeignKey('clients.id'), nullable=False)
    template_id = db.Column(db.Integer, db.ForeignKey('contract_templates.id'))
    # Lifecycle: draft → in_review → approved → finalized → partially_executed → fully_executed
    # expired is a terminal state reachable from any active status
    status = db.Column(db.String(50), default='draft')
    notes = db.Column(db.Text)
    start_date = db.Column(db.Date)
    end_date = db.Column(db.Date)
    value = db.Column(db.Numeric(15, 2))
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    finalized_at = db.Column(db.DateTime)   # when sent for signature
    executed_at  = db.Column(db.DateTime)   # when fully executed (all parties signed)

    revisions = db.relationship('ContractRevision', backref='contract', lazy='dynamic', order_by='ContractRevision.version_number')
    field_values = db.relationship('ContractFieldValue', backref='contract', lazy='dynamic')
    audit_logs = db.relationship('AuditLog', backref='contract', lazy='dynamic')

    @property
    def latest_revision(self):
        return self.revisions.order_by(ContractRevision.version_number.desc()).first()

    @property
    def revision_count(self):
        return self.revisions.count()

    @property
    def is_executed(self):
        return self.status == 'fully_executed'

    @property
    def status_color(self):
        colors = {
            'draft':              'gray',
            'in_review':          'yellow',
            'approved':           'blue',
            'finalized':          'purple',
            'partially_executed': 'orange',
            'fully_executed':     'green',
            'expired':            'red',
        }
        return colors.get(self.status, 'gray')

    def __repr__(self):
        return f'<Contract {self.title}>'


class ContractRevision(db.Model):
    __tablename__ = 'contract_revisions'
    id = db.Column(db.Integer, primary_key=True)
    contract_id = db.Column(db.Integer, db.ForeignKey('contracts.id'), nullable=False)
    version_number = db.Column(db.Integer, nullable=False)
    content = db.Column(db.Text, nullable=False)
    changes_summary = db.Column(db.Text)
    file_path = db.Column(db.String(500))
    file_name = db.Column(db.String(255))
    is_finalized = db.Column(db.Boolean, default=False)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    creator = db.relationship('User', foreign_keys=[created_by])

    def __repr__(self):
        return f'<ContractRevision {self.contract_id} v{self.version_number}>'


class ContractFieldValue(db.Model):
    __tablename__ = 'contract_field_values'
    id = db.Column(db.Integer, primary_key=True)
    contract_id = db.Column(db.Integer, db.ForeignKey('contracts.id'), nullable=False)
    field_name = db.Column(db.String(255), nullable=False)
    field_value = db.Column(db.Text)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def __repr__(self):
        return f'<ContractFieldValue {self.field_name}>'


class AuditLog(db.Model):
    __tablename__ = 'audit_logs'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    contract_id = db.Column(db.Integer, db.ForeignKey('contracts.id'))
    action = db.Column(db.String(100), nullable=False)
    resource_type = db.Column(db.String(100))
    resource_id = db.Column(db.Integer)
    details = db.Column(db.Text)
    ip_address = db.Column(db.String(50))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f'<AuditLog {self.action} at {self.created_at}>'
