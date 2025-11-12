from flask import Flask, render_template, request, jsonify
import os
from werkzeug.utils import secure_filename
from utils.excel_processor import processar_capacidade
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from datetime import datetime
from dotenv import load_dotenv
import base64

load_dotenv()

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# Criar pasta de uploads se não existir
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'xlsm'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/processar', methods=['POST'])
def processar():
    try:
        # Verificar se arquivo foi enviado
        if 'arquivo' not in request.files:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        file = request.files['arquivo']
        dia = request.form.get('dia')
        
        if file.filename == '':
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        if not dia:
            return jsonify({'error': 'Selecione o dia'}), 400
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Processar o arquivo Excel
            dados = processar_capacidade(filepath, int(dia))
            
            # Remover arquivo após processamento
            os.remove(filepath)
            
            return jsonify({'success': True, 'dados': dados})
        else:
            return jsonify({'error': 'Formato de arquivo não permitido. Use .xlsx, .xls ou .xlsm'}), 400
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/enviar_email', methods=['POST'])
def enviar_email():
    try:
        # Receber a imagem em base64
        data = request.get_json()
        image_data = data.get('image')
        
        if not image_data:
            return jsonify({'error': 'Imagem não fornecida'}), 400
        
        # Decodificar base64
        image_bytes = base64.b64decode(image_data.split(',')[1])
        
        # Configurar destinatários
        email_to = [
            'Diego.Appolari@unilever.com',
            'Andre.Martinato@unilever.com',
            'Guilherme.Frigo@unilever.com',
            'Cristiano.Felten@unilever.com'
        ]
        email_cc = [
            'Michel.Lopes@unilever.com','Edson.Filgueiras@unilever.com','Merquides.Guimaraes@unilever.com',
            'Gilson.Lima@unilever.com','Jonas.Olivatto@unilever.com','Evelyn.Missio@unilever.com',
            'Wagner.Correa@unilever.com','Thiago.Roque@unilever.com','Clayton.Scodeler@unilever.com',
            'Simoni.Serafim@unilever.com','Geise.Silva@unilever.com','Telma.Silva@unilever.com',
            'Diego.Soaress@unilever.com','Kananda.Gouvea@unilever.com','Cleber.Rizzo@unilever.com',
            'Denicielle.Otaviano@unilever.com','Leandro.Neves@unilever.com','Marcos.Felix@unilever.com',
            'Claudio.Marques@unilever.com','WENDELL.COSTA-DUARTE@unilever.com','Diego.Sousa@unilever.com',
            'Renato.Segli@unilever.com','Brunna.Arruda@unilever.com','Carolina.Ouno@unilever.com',
            'Pedro.Nobre@unilever.com','Celso.Leitao@unilever.com','Felipe.Lobo@unilever.com',
            'Renata.Azevedo@unilever.com','Rafael.Ribeiro@unilever.com','Tamires.Turatti@unilever.com'
        ]
        
        # Saudação e data
        hora = datetime.now().hour
        if hora < 12:
            saudacao = 'Bom dia'
        elif hora < 18:
            saudacao = 'Boa tarde'
        else:
            saudacao = 'Boa noite'
        
        data_hoje = datetime.now().strftime('%d/%m/%Y')
        assunto = f'Update Dock {data_hoje}'
        
        # Criar mensagem HTML com imagem inline
        html_body = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
            <p>{saudacao},</p>
            <p>Segue a informação do dock:</p>
            <br>
            <img src="cid:dock_image" style="max-width: 100%; height: auto;">
        </body>
        </html>
        """
        
        # Criar mensagem multipart
        msg = MIMEMultipart('related')
        msg['Subject'] = assunto
        msg['From'] = os.getenv('EMAIL_USER', 'seu-email@unilever.com')
        msg['To'] = ', '.join(email_to)
        msg['Cc'] = ', '.join(email_cc)
        
        # Adicionar corpo HTML
        msg_alternative = MIMEMultipart('alternative')
        msg.attach(msg_alternative)
        msg_alternative.attach(MIMEText(html_body, 'html'))
        
        # Adicionar imagem inline
        img = MIMEImage(image_bytes, name='dock_update.png')
        img.add_header('Content-ID', '<dock_image>')
        img.add_header('Content-Disposition', 'inline', filename='dock_update.png')
        msg.attach(img)
        
        # Enviar via SMTP
        smtp_server = os.getenv('SMTP_SERVER', 'smtp.office365.com')
        smtp_port = int(os.getenv('SMTP_PORT', '587'))
        email_user = os.getenv('EMAIL_USER')
        email_pass = os.getenv('EMAIL_PASS')
        
        if not email_user or not email_pass:
            return jsonify({'error': 'Credenciais de e-mail não configuradas no servidor'}), 500
        
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(email_user, email_pass)
            recipients = email_to + email_cc
            server.send_message(msg, to_addrs=recipients)
        
        return jsonify({'success': True, 'message': 'E-mail enviado com sucesso!'})
        
    except Exception as e:
        return jsonify({'error': f'Erro ao enviar e-mail: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
