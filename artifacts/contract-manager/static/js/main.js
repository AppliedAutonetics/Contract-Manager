// Flash message auto-dismiss
document.querySelectorAll('.flash-close').forEach(btn => {
  btn.addEventListener('click', () => btn.closest('.flash').remove());
});
setTimeout(() => {
  document.querySelectorAll('.flash').forEach(f => {
    f.style.transition = 'opacity 0.4s, transform 0.4s';
    f.style.opacity = '0';
    f.style.transform = 'translateX(20px)';
    setTimeout(() => f.remove(), 400);
  });
}, 4500);

// File drag-and-drop
document.querySelectorAll('.file-upload-area').forEach(area => {
  const input = area.querySelector('input[type="file"]');
  area.addEventListener('click', () => input.click());
  area.addEventListener('dragover', e => { e.preventDefault(); area.classList.add('drag-over'); });
  area.addEventListener('dragleave', () => area.classList.remove('drag-over'));
  area.addEventListener('drop', e => {
    e.preventDefault();
    area.classList.remove('drag-over');
    if (e.dataTransfer.files.length) {
      input.files = e.dataTransfer.files;
      updateFileLabel(area, e.dataTransfer.files[0].name);
    }
  });
  if (input) {
    input.addEventListener('change', () => {
      if (input.files.length) updateFileLabel(area, input.files[0].name);
    });
  }
});

function updateFileLabel(area, name) {
  const text = area.querySelector('.upload-text');
  if (text) text.textContent = '📄 ' + name;
}

// Template field detection
const templateSelect = document.getElementById('template_id');
const fieldsContainer = document.getElementById('template-fields-container');

if (templateSelect && fieldsContainer) {
  templateSelect.addEventListener('change', () => {
    const tid = templateSelect.value;
    if (!tid) { fieldsContainer.innerHTML = ''; return; }
    fetch(`/api/templates/${tid}/fields`)
      .then(r => r.json())
      .then(data => {
        if (!data.fields || data.fields.length === 0) {
          fieldsContainer.innerHTML = '<p class="text-sm text-muted">No fillable fields detected in this template.</p>';
          return;
        }
        let html = '<div class="card mt-4"><div class="card-header"><span class="card-title">Template Fields</span><span class="text-sm text-muted">' + data.fields.length + ' fields found</span></div><div class="card-body"><div class="fields-list mb-4">';
        data.fields.forEach(f => {
          html += `<span class="field-tag">{{${f}}}</span>`;
        });
        html += '</div><div class="form-row" id="fields-grid">';
        data.fields.forEach(f => {
          const label = f.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
          html += `<div class="form-group"><label class="form-label">${label}</label><input type="text" name="field_${f}" class="form-control" placeholder="Enter value for {{${f}}}"></div>`;
        });
        html += '</div></div></div>';
        fieldsContainer.innerHTML = html;
      })
      .catch(console.error);
  });
}

// Live field detection in template editor
const contentTextarea = document.getElementById('template-content');
const detectedFields = document.getElementById('detected-fields');

if (contentTextarea && detectedFields) {
  let debounceTimer;
  contentTextarea.addEventListener('input', () => {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => {
      const content = contentTextarea.value;
      const fields = [...new Set([...content.matchAll(/\{\{([A-Z_][A-Z0-9_]*)\}\}/g)].map(m => m[1]))];
      if (fields.length === 0) {
        detectedFields.innerHTML = '<span class="text-muted text-sm">No fields detected. Use {{FIELD_NAME}} syntax (uppercase).</span>';
      } else {
        detectedFields.innerHTML = fields.map(f => `<span class="field-tag">{{${f}}}</span>`).join('');
      }
    }, 300);
  });
  contentTextarea.dispatchEvent(new Event('input'));
}

// Confirm delete/finalize
document.querySelectorAll('[data-confirm]').forEach(el => {
  el.addEventListener('click', e => {
    if (!confirm(el.dataset.confirm)) e.preventDefault();
  });
});

// Revision compare form auto-redirect
const compareForm = document.getElementById('compare-form');
if (compareForm) {
  compareForm.querySelectorAll('select').forEach(sel => {
    sel.addEventListener('change', () => compareForm.submit());
  });
}

// Active nav item
const currentPath = window.location.pathname;
document.querySelectorAll('.nav-item').forEach(item => {
  const href = item.getAttribute('href');
  if (!href) return;
  if (currentPath === href || (href !== '/' && currentPath.startsWith(href))) {
    item.classList.add('active');
  }
});
