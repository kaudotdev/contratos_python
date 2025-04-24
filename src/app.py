from flask import Flask, render_template, request, send_file, redirect, url_for
from docx import Document
import os
from datetime import datetime
import io

app = Flask(__name__)
UPLOAD_FOLDER = 'templates'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def fill_contract(template_path, data):
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        paragraph_text = paragraph.text
        for key, value in data.items():
            if key in paragraph_text:
                paragraph_text = paragraph_text.replace(key, value)

        if paragraph.text != paragraph_text:
            for run in paragraph.runs:
                run.text = ""
            paragraph.runs[0].text = paragraph_text

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    return file_stream


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/form', methods=['GET', 'POST'])
def form():
    if request.method == 'POST':
        contract_type = request.form.get('contract_type')
        return render_template('contract_form.html', contract_type=contract_type)
    return render_template('select_type.html')


@app.route('/generate', methods=['POST'])
def generate_contract():
    contract_type = request.form.get('contract_type')

    data = {
        '[unidade]': request.form.get('unidade', ''),
        '[tipo]': request.form.get('tipo', ''),
        '[metro]': request.form.get('metro', ''),
        '[metrext]': request.form.get('metrext', ''),
        '[preço]': request.form.get('preco', ''),
        '[preçoext]': request.form.get('precoext', ''),
        '[nome1]': request.form.get('nome1', ''),
        '[nac1]': request.form.get('nac1', ''),
        '[prof1]': request.form.get('prof1', ''),
        '[cpf1]': request.form.get('cpf1', ''),
        '[rg1]': request.form.get('rg1', ''),
        '[tel1]': request.form.get('tel1', ''),
        '[email1]': request.form.get('email1', ''),
        '[data]': request.form.get('data', datetime.now().strftime("%d/%m/%Y"))
    }

    template_path = ""

    if contract_type == "casados":
        template_path = "templates/Contrato dois mutuantes casados.docx"
        data['[end]'] = request.form.get('end', '')
        data['[nome2]'] = request.form.get('nome2', '')
        data['[nac2]'] = request.form.get('nac2', '')
        data['[prof2]'] = request.form.get('prof2', '')
        data['[cpf2]'] = request.form.get('cpf2', '')
        data['[rg2]'] = request.form.get('rg2', '')
        data['[tel2]'] = request.form.get('tel2', '')
        data['[email2]'] = request.form.get('email2', '')

    elif contract_type == "nao_casados":
        template_path = "templates/Contrato dois mutuantes não casados.docx"
        data['[end1]'] = request.form.get('end1', '')
        data['[end2]'] = request.form.get('end2', '')
        data['[nome2]'] = request.form.get('nome2', '')
        data['[nac2]'] = request.form.get('nac2', '')
        data['[prof2]'] = request.form.get('prof2', '')
        data['[cpf2]'] = request.form.get('cpf2', '')
        data['[rg2]'] = request.form.get('rg2', '')
        data['[tel2]'] = request.form.get('tel2', '')
        data['[email2]'] = request.form.get('email2', '')
        data['[ec1]'] = request.form.get('ec1', '')
        data['[ec2]'] = request.form.get('ec2', '')

    elif contract_type == "solteiro":
        template_path = "templates/Contrato mutuante solteiro.docx"
        data['[end]'] = request.form.get('end', '')
        data['[ec1]'] = request.form.get('ec1', '')

    if os.path.exists(template_path):
        file_stream = fill_contract(template_path, data)

        return send_file(
            file_stream,
            as_attachment=True,
            download_name=f"{data['[unidade]']}.BRID-Contrato de mútuo.docx",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        return "Erro: Modelo de contrato não encontrado", 404


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
