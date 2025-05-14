document.getElementById('load').onclick = async () => {
  const id = document.getElementById('rowId').value;
  const res = await fetch(`/api/row/${id}`);
  if (!res.ok) return alert('No existe esa fila');
  const data = await res.json();
  const container = document.getElementById('fields');
  container.innerHTML = '';
  for (let k in data) {
    if (k === 'id') continue;
    const div = document.createElement('div');
    div.innerHTML = `<label>${k}: <input name="${k}" value="${data[k]}"></label>`;
    container.appendChild(div);
  }
  document.getElementById('editForm').style.display = 'block';
  document.getElementById('editForm').onsubmit = async e => {
    e.preventDefault();
    const formData = Object.fromEntries(new FormData(e.target).entries());
    const upd = await fetch(`/api/row/${id}`, {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify(formData)
    });
    if (upd.ok) {
      window.open(`/api/row/${id}/pdf`, '_blank');
    }
  };
};
