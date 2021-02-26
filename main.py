# coding: utf-8
import xml.etree.ElementTree as ET
import os
import glob

prefixo_tag_xml = '{http://www.portalfiscal.inf.br/nfe}'
pasta_raiz = r'D:\HENRIQUE\PYTHON\Programas\XML to Excel\NOTAS DE SAIDA'


def numero_nota_fiscal(xml, prefixo=prefixo_tag_xml):    
    # retorna o numero da nota fiscal
    for numero_nota in xml.iter(prefixo + 'nNF'):
        return numero_nota.text

def data_emissao_nota_fiscal(xml, prefixo=prefixo_tag_xml):
    # retorna a data de emissao da nota fiscal
    for data_emissao in xml.iter(prefixo + 'dhEmi'):
        return data_emissao.text[:10]

def tipo_operacao_nota_fiscal(xml, prefixo=prefixo_tag_xml):
    # retorna se a nota fiscal é de entrada ou saida
    # 0 = entrada
    # 1 = saida
    for tipo_operacao in xml.iter(prefixo + 'tpNF'):        
        return tipo_operacao.text

def operacao_consumidor_final(xml, prefixo=prefixo_tag_xml):
    # retorna se a operação é destinada a consumidor final
    # 0 = normal
    # 1 = consumidor final
    for operacao_consumidor in xml.iter(prefixo + 'indFinal'):
        return operacao_consumidor.text

def cnpj_cpf_identificacao_destinatario_nf(xml, prefixo=prefixo_tag_xml):
    # retorna o CPF ou CNPJ  do destinatario da nf
    for destinatario in xml.iter(prefixo + 'dest'):
        for identificacao in destinatario:
            return identificacao.text

def nome_identificacao_destinatario(xml, prefixo=prefixo_tag_xml):
    # retorna o nome do destinatario
    for nome_destinatario in xml.iter(prefixo + 'dest'):
        for nome in nome_destinatario.iter(prefixo + 'xNome'):
            return nome.text

def estado_uf_destinatario_nf(xml, prefixo=prefixo_tag_xml):
    # retorna o estado do destinatario
    for estado_uf in xml.iter(prefixo + 'dest'):
        for uf in estado_uf.iter(prefixo + 'UF'):
            return uf.text
        
def dados_produtos_nota_fiscal(xml, prefixo=prefixo_tag_xml):
    # retorna as informações de todos os produtos da nota fiscal
    # numero do item ; codigo do produto ; nome do produto; NCM; CFOP;
    # quantidade comercializada; valor do unitario comercializado ; valor total
    lista_itens = []
    
    for item in xml.iter(prefixo + 'det'):
        numero_item = item.get('nItem')
            
        for produto in item.iter(prefixo + 'prod'):
            cod_produto = produto.find(prefixo + 'cProd').text
            nome_produto = produto.find(prefixo + 'xProd').text
            ncm_produto = produto.find(prefixo + 'NCM').text
            cfop_produto = produto.find(prefixo + 'CFOP').text
            qtd_comercializada = produto.find(prefixo + 'qCom').text
            valor_un_comercializado = produto.find(prefixo + 'vUnCom').text
            valor_total_prod = produto.find(prefixo + 'vProd').text


        lista_itens.append([
            numero_item,
            cod_produto,
            nome_produto,
            ncm_produto,
            cfop_produto,
            qtd_comercializada,
            valor_un_comercializado,
            valor_total_prod
            ])

    return lista_itens
  


def rodar_listas_xml(nome_arquivo_txt='resultado'):    
    lista_arquivos = glob.glob('*.xml')
    quantidade_xml = len(glob.glob('*.xml'))

    with open(nome_arquivo_txt, 'w') as arquivo_txt:
        arquivo_txt.write('NUMERO NOTA;DATA EMISSAO;TIPO DE OPERAÇÃO;CONSUMIDOR FINAL;CNPJ/CPF;RAZÃO SOCIAL;UF;ITEM NA NOTA;CODIGO DO PRODUTO;DESCRIÇÃO;NCM;CFOP;QUANTIDADE;VALOR UNITARIO;VALOR TOTAL\n')
        try:
            for posicao in range(quantidade_xml):
                xml_nota_fiscal = ET.parse(lista_arquivos[posicao]).getroot()
                
                numero_nota = numero_nota_fiscal(xml_nota_fiscal)
                data_emissao_nota = data_emissao_nota_fiscal(xml_nota_fiscal)
                tipo_operacao = tipo_operacao_nota_fiscal(xml_nota_fiscal)
                consumidor_final = operacao_consumidor_final(xml_nota_fiscal)
                cpf_cnpj_destinatario = cnpj_cpf_identificacao_destinatario_nf(xml_nota_fiscal)
                nome_destinatario = nome_identificacao_destinatario(xml_nota_fiscal)
                estado_destinatario = estado_uf_destinatario_nf(xml_nota_fiscal)
                dados_produtos = dados_produtos_nota_fiscal(xml_nota_fiscal)

                for itens in dados_produtos:
                    arquivo_txt.write(
                        numero_nota + ';' +
                        data_emissao_nota + ';' +
                        tipo_operacao + ';' +
                        consumidor_final + ';' +
                        cpf_cnpj_destinatario + ';' +
                        nome_destinatario + ';' +
                        estado_destinatario + ';')
                    
                    for item in itens:
                        arquivo_txt.write(
                            item + ';')
                    arquivo_txt.write('\n')
        except:
            print('Erro')


rodar_listas_xml()
print('Concluído')







