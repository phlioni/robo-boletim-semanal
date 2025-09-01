# -*- coding: utf-8 -*-
from azure.devops.connection import Connection
from msrest.authentication import BasicAuthentication
import config

_connections = {} # Cache para evitar logins repetidos na mesma organização

def _get_connection(org_url):
    """Cria e armazena em cache uma conexão para uma organização do Azure DevOps."""
    if org_url not in _connections:
        print(f"Criando nova conexão para a organização: {org_url}")
        credentials = BasicAuthentication('', config.AZURE_DEVOPS_PAT)
        _connections[org_url] = Connection(base_url=org_url, creds=credentials)
    return _connections[org_url]

def get_work_items_details(org_url, project_name, work_item_ids):
    """Busca os detalhes completos de uma lista de Work Items."""
    if not work_item_ids:
        return {}

    # Remove duplicados e converte para lista de inteiros
    ids_unicos = list(set(int(wid) for wid in work_item_ids if wid is not None))
    if not ids_unicos:
        return {}

    try:
        connection = _get_connection(org_url)
        wit_client = connection.clients.get_work_item_tracking_client()
        
        print(f"Buscando {len(ids_unicos)} Work Items em '{project_name}' no Azure DevOps...")
        
        # A API tem um limite de 200 IDs por chamada
        work_items_details = wit_client.get_work_items(ids=ids_unicos[:200], expand="All", error_policy="omit")
        
        dados_devops = {}
        for item in work_items_details:
            dados_devops[item.id] = {
                "titulo": item.fields.get("System.Title", "N/A"),
                "status": item.fields.get("System.State", "N/A"),
                "tipo": item.fields.get("System.WorkItemType", "N/A"),
                "descricao_html": item.fields.get("System.Description", "") or item.fields.get("Microsoft.VSTS.TCM.ReproSteps", ""),
                "tags": item.fields.get("System.Tags", ""),
                "estimativa_original": item.fields.get("Microsoft.VSTS.Scheduling.OriginalEstimate", 0),
                "esforco_concluido": item.fields.get("Microsoft.VSTS.Scheduling.CompletedWork", 0)
            }
        
        print(f" -> {len(dados_devops)} detalhes de Work Items encontrados.")
        return dados_devops

    except Exception as e:
        print(f"ERRO ao buscar detalhes de Work Items em '{project_name}': {e}")
        return {}