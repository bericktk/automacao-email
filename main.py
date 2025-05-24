import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import os
import pandas as pd
from email.utils import formataddr

try:
    from dados import *
except ImportError:
    print("!! ERRO: Arquivo 'dados.py' não encontrado ou variáveis não definidas nele.")
    print("   Por favor, crie o arquivo senha.py com o conteúdo: nomeDoRemetente, emailDoRemetente, senhaAppGoogle, servidorDoSMTP, portaDoSMTP")
    senhaAppGoogle = "SUA_SENHA_DE_APP_NAO_CONFIGURADA"

nomeRemetente = nomeDoRemetente
emailRemetente = emailDoRemetente
senhaAppGmail = senhaAppGoogle
servidorSMTP = servidorDoSMTP
portaSMTP = portaDoSMTP

imagemAssinaturaEmail = 'assinatura_Bruno_Erick.png'

assinaturaEmail = f"""
<p>Atenciosamente,<br>
<strong>Bruno Erick</strong><br>
+55 85 99140-6794</p><br>
<img src="cid:imagemAssinaturaEmail" alt="Assinatura Bruno Erick" style="max-width: 100%; height: auto;">
"""

arquivoPlanilhaFaturas = 'clientes_cbpce.xlsx'
nomeAbaPlanilha = 'Plan1'
planilhaFalhas = 'log_falhas_envios.xlsx'

def carregar_dados_faturas(caminho_arquivo, nome_aba):
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name=nome_aba)
        colunas_obrigatorias = ['Nome', 'Email', 'ArquivoFatura', 'Valor', 'Vencimento']
        for coluna in colunas_obrigatorias:
            if coluna not in df.columns:
                print(f"!! ERRO: Coluna obrigatória '{coluna}' não encontrada na planilha '{caminho_arquivo}', aba '{nome_aba}'.")
                print(f"   Colunas encontradas: {list(df.columns)}")
                return None
        print(f"Dados carregados com sucesso da planilha: {caminho_arquivo} (Aba: {nome_aba})")
        return df
    except FileNotFoundError:
        print(f"!! ERRO: Arquivo da planilha não encontrado: {caminho_arquivo}")
        return None
    except Exception as e:
        print(f"!! ERRO ao ler a planilha '{caminho_arquivo}': {e}")
        return None

def enviar_email_com_fatura(destinatario_email, destinatario_nome, emails_cc, assunto_email, corpo_html_email, caminho_anexo_fatura, caminho_imagem_assinatura_local):
    try:
        msg_raiz = MIMEMultipart('related')
        msg_raiz['From'] = formataddr((nomeRemetente, emailRemetente))
        msg_raiz['To'] = destinatario_email
        msg_raiz['Subject'] = assunto_email

        lista_todos_destinatarios = [destinatario_email]

        if emails_cc:
            emails_cc_validos = [email.strip() for email in emails_cc if email and email.strip()]
            if emails_cc_validos:
                msg_raiz['Cc'] = ", ".join(emails_cc_validos)
                lista_todos_destinatarios.extend(emails_cc_validos)

        msg_alternativa = MIMEMultipart('alternative')
        msg_raiz.attach(msg_alternativa)

        parte_html = MIMEText(corpo_html_email, 'html')
        msg_alternativa.attach(parte_html)

        if os.path.exists(caminho_imagem_assinatura_local):
            try:
                with open(caminho_imagem_assinatura_local, 'rb') as fp_imagem:
                    img = MIMEImage(fp_imagem.read())
                img.add_header('Content-ID', '<imagemAssinaturaEmail>')
                msg_raiz.attach(img)
            except Exception as e_img:
                print(f"!! AVISO: Não foi possível incorporar a imagem da assinatura '{caminho_imagem_assinatura_local}': {e_img}")
        else:
            print(f"!! AVISO: Ficheiro da imagem da assinatura não encontrado em '{caminho_imagem_assinatura_local}'. A assinatura será enviada sem imagem.")

        if not os.path.exists(caminho_anexo_fatura):
            erro_msg = f"Arquivo de anexo (fatura) não encontrado: {caminho_anexo_fatura}"
            print(f"!! ERRO: {erro_msg} para {destinatario_nome}")
            return False, erro_msg

        nome_arquivo_anexo_fatura = os.path.basename(caminho_anexo_fatura)
        with open(caminho_anexo_fatura, 'rb') as anexo_fatura_fp:
            parte_anexo_fatura = MIMEBase('application', 'octet-stream')
            parte_anexo_fatura.set_payload(anexo_fatura_fp.read())
        encoders.encode_base64(parte_anexo_fatura)
        parte_anexo_fatura.add_header(
            'Content-Disposition',
            f'attachment; filename="{nome_arquivo_anexo_fatura}"',
        )
        msg_raiz.attach(parte_anexo_fatura)

        print(f"-> Conectando ao servidor SMTP para enviar para: {destinatario_nome} ({destinatario_email}) CC: {emails_cc if emails_cc else 'Nenhum'}...")
        servidor = smtplib.SMTP(servidorSMTP, portaSMTP)
        servidor.ehlo()
        servidor.starttls()
        servidor.ehlo()
        servidor.login(emailRemetente, senhaAppGmail)
        print("   Conectado e logado com sucesso!")
        texto_completo_email = msg_raiz.as_string()
        servidor.sendmail(emailRemetente, lista_todos_destinatarios, texto_completo_email)
        print(f"   E-mail para {destinatario_nome} enviado com sucesso!")
        servidor.quit()
        return True, None

    except FileNotFoundError as e:
        erro_msg = f"Erro de ficheiro não encontrado: {e}"
        print(f"!! ERRO CRÍTICO: {erro_msg} para {destinatario_nome}")
        return False, erro_msg
    except smtplib.SMTPAuthenticationError as e:
        erro_msg = f"ERRO DE AUTENTICAÇÃO SMTP: Verifique seu e-mail ({emailRemetente}) e senha de app. - {e}"
        print(f"!! {erro_msg} Falha para {destinatario_nome}.")
        return False, erro_msg
    except Exception as e:
        erro_msg = f"Erro ao enviar e-mail: {e}"
        print(f"!! {erro_msg} para {destinatario_nome} ({destinatario_email})")
        if 'servidor' in locals() and servidor:
            try:
                servidor.quit()
            except Exception:
                pass
        return False, erro_msg

if __name__ == "__main__":
    print("--- Iniciando processo de envio de faturas ---")

    if senhaAppGmail == "SUA_SENHA_DE_APP_NAO_CONFIGURADA":
        print("!! ATENÇÃO: A senha de app não foi configurada corretamente. Verifique o arquivo 'senha.py'.")
        print("   O script não prosseguirá sem a configuração da senha.")
    else:
        emails_enviados_com_sucesso = 0
        emails_com_falha = 0
        lista_falhas_envio = []

        if not os.path.exists(imagemAssinaturaEmail):
            print(f"!!! ATENÇÃO: Ficheiro da imagem da assinatura '{imagemAssinaturaEmail}' não encontrado. Os e-mails serão enviados sem a imagem na assinatura.")

        clientes_faturas_df = carregar_dados_faturas(arquivoPlanilhaFaturas, nomeAbaPlanilha)

        if clientes_faturas_df is not None:
            pasta_base_faturas = 'faturas_pdf'
            if not os.path.exists(pasta_base_faturas):
                os.makedirs(pasta_base_faturas)
                print(f"Pasta '{pasta_base_faturas}' criada. Certifique-se que os caminhos na planilha apontam para cá ou para os locais corretos.")

            for indice, linha in clientes_faturas_df.iterrows():
                nome_cli = linha['Nome']
                email_cli = linha['Email']
                fatura_path_cli = linha['ArquivoFatura']
                vencimento_cli = linha['Vencimento']
                
                emails_cc_lista = []
                if 'EmailCopia' in linha and pd.notna(linha['EmailCopia']):
                    emails_cc_str = str(linha['EmailCopia'])
                    emails_cc_lista = [email.strip() for email in emails_cc_str.replace(';', ',').split(',') if email.strip()]
                
                print(f"\nProcessando cliente: {nome_cli} (Linha {indice + 2} da planilha)")

                assunto = f"Boleto Mensalidade CBPCE Jun25 - {nome_cli}"
                corpo_html = f"""
                <html>
                <head>
                    <style>
                        body {{ font-family: Arial, sans-serif; line-height: 1.6; max-width: 100%; }}
                        .container {{ padding: 20px; border: 1px solid #ddd; border-radius: 5px; max-width: 600px; margin: 20px auto; }}
                        .header {{ font-size: 1.2em; font-weight: bold; color: #333; }}
                    </style>
                </head>
                <body>
                    <div class="container">
                        <p class="header">Olá prezados, como estão?</p>
                        <p>Encaminhamos em anexo o boleto da {nome_cli} referente a mensalidade associativa da CBPCE com vencimento para o dia {vencimento_cli}.</p>
                        <p>Caso não consiga realizar o pagamento via boleto, segue abaixo os nossos dados bancários para depósito.</p>
                        <p><strong>Banco do Brasil</strong><br>
                        Ag. 2917-3<br>
                        Conta 70077-0<br>
                        CNPJ 04.549.837/0001-45<br>
                        Chave Pix: 04549837000145
                        </p>
                        <p>Por favor acusar o recebimento, em caso de dúvidas entre em contato conosco.</p>
                        {assinaturaEmail}
                    </div>
                </body>
                </html>
                """

                sucesso_envio, mensagem_erro = enviar_email_com_fatura(email_cli, nome_cli, emails_cc_lista, assunto, corpo_html, fatura_path_cli, imagemAssinaturaEmail)

                if sucesso_envio:
                    emails_enviados_com_sucesso += 1
                else:
                    emails_com_falha += 1
                    lista_falhas_envio.append({
                        'NomeCliente': nome_cli,
                        'EmailPrincipal': email_cli,
                        'EmailsCopia': ", ".join(emails_cc_lista) if emails_cc_lista else '',
                        'ArquivoFatura': fatura_path_cli,
                        'HorarioFalha': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'MotivoFalha': mensagem_erro if mensagem_erro else 'Erro desconhecido durante o envio.'
                    })
        else:
            print("!! Processo de envio de faturas não pode continuar devido a erro no carregamento da planilha.")

        if lista_falhas_envio:
            print(f"\n--- Registrando {len(lista_falhas_envio)} falha(s) de envio no arquivo: {planilhaFalhas} ---")
            df_falhas = pd.DataFrame(lista_falhas_envio)
            try:
                df_falhas.to_excel(planilhaFalhas, index=False, sheet_name='FalhasEnvio')
                print(f"Arquivo de log de falhas '{planilhaFalhas}' criado com sucesso.")
            except Exception as e:
                print(f"!! ERRO ao criar o arquivo de log de falhas '{planilhaFalhas}': {e}")
        else:
            if clientes_faturas_df is not None:
                 print("\nNenhuma falha de envio registrada.")

        print("\n--- Processo de envio de faturas concluído ---")
        print(f"Total de e-mails enviados com sucesso: {emails_enviados_com_sucesso}")
        print(f"Total de e-mails com falha no envio: {emails_com_falha}")