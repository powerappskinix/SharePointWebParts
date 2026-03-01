import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ComunicadosWebPart.module.scss';

export interface IComunicadosProps {
  FundoPainel: string;
  AreasExibicao: string; // Nova propriedade para o filtro
}

interface IComunicadoItem {
  Id: number;
  Title: string;
  Descricao: string; 
  CorCard: string;
  Icone: string;
  StatusComunicado: string;
  CorTexto: string; 
  Area: string; // Corresponde à coluna Choice do SharePoint
}

export default class ComunicadosWebPart extends BaseClientSideWebPart<IComunicadosProps> {
  private _sp: SPFI;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
  }

  public render(): void {
    this._obterComunicados().then(itens => {
      const valorFundo = this.properties.FundoPainel || "#f3f2f1";
      
      const estiloFundo = valorFundo.indexOf('http') > -1 
        ? `background-image: url('${valorFundo}'); background-size: cover; background-position: center;` 
        : `background-color: ${valorFundo};`;

      const cardsHtml = itens.length > 0 
        ? itens.map(item => this._renderCard(item)).join('')
        : `<p style="color:#666; font-family:'Asap'; text-align:center; width:100%; padding: 20px;">Nenhum comunicado disponível para as áreas selecionadas.</p>`;

      this.domElement.innerHTML = `
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Asap:wght@400;700&display=swap');
          .pin-visual {
            width: 14px;
            height: 14px;
            background-color: #e63946;
            border-radius: 50%;
            position: absolute;
            top: 10px;
            left: 50%;
            transform: translateX(-50%);
            box-shadow: 0 2px 3px rgba(0,0,0,0.3);
            z-index: 5;
          }
        </style>
        
        <div class="${styles.comunicados}" style="
          ${estiloFundo}
          padding: 30px 20px;
          border-radius: 8px;
          min-height: auto;
          display: flow-root;
        ">
          <div style="display: flex; flex-wrap: wrap; gap: 20px; justify-content: flex-start; max-width: 1400px; margin: 0 auto;">
            ${cardsHtml}
          </div>
        </div>
      `;
    });
  }

  private _renderCard(item: IComunicadoItem): string {
    const corFundo = item.CorCard || "#f4f4f4";
    const corDoTextoDinamica = item.CorTexto || "#ffffff"; 
    const temIcone = item.Icone && item.Icone.trim() !== "";
    
    return `
      <div style="
        background-color: ${corFundo};
        flex: 0 1 280px; 
        min-height: 200px;
        padding: 35px 20px 25px 20px;
        border-radius: 4px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        position: relative;
        font-family: 'Asap', sans-serif;
        display: flex;
        flex-direction: column;
      ">
        <div class="pin-visual"></div>
        
        <div style="display: flex; align-items: center; width: 100%; margin-bottom: 10px;">
          <div style="width: 30px; flex-shrink: 0; font-size: 20px; display: flex; align-items: center; justify-content: flex-start;">
            ${temIcone ? item.Icone : ''}
          </div>
          <div style="flex-grow: 1; display: flex; justify-content: center; align-items: center;">
            <span style="font-size: 19px; font-weight: 700; color: ${corDoTextoDinamica}; text-transform: uppercase; line-height: 1.1; text-align: center;">
              ${escape(item.Title)}
            </span>
          </div>
          <div style="width: 30px; flex-shrink: 0;"></div>
        </div>
        
        <div style="width: 100%; border-top: 1px solid ${corDoTextoDinamica}; opacity: 0.3; margin-bottom: 20px;"></div>

        <div style="width: 100%; flex-grow: 1; display: flex; align-items: center; justify-content: center;">
          <div style="font-size: 20px; font-weight: 400; line-height: 1.4; color: ${corDoTextoDinamica}; ">
            ${escape(item.Descricao)}
          </div>
        </div>
      </div>
    `;
  }

  private async _obterComunicados(): Promise<IComunicadoItem[]> {
    try {
      let queryFiltro = "StatusComunicado eq 'Ativo'";

      // Adiciona filtro por Área se preenchido no Painel de Propriedades
      if (this.properties.AreasExibicao) {
        const areas = this.properties.AreasExibicao.split(',')
          .map(a => a.trim())
          .filter(a => a.length > 0);

        if (areas.length > 0) {
          const filtroAreas = areas.map(a => `Area eq '${a}'`).join(' or ');
          queryFiltro += ` and (${filtroAreas})`;
        }
      }

      return await this._sp.web.lists.getByTitle("Comunicados").items
        .select("Id", "Title", "Descricao", "CorCard", "Icone", "StatusComunicado", "CorTexto", "Area")
        .filter(queryFiltro)();
    } catch (error) {
      console.error("Erro ao obter lista:", error);
      return [];
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        groups: [{
          groupName: "Design do Quadro",
          groupFields: [
            PropertyPaneTextField('FundoPainel', { 
              label: 'Fundo do Quadro',
              description: 'URL de imagem ou Cor Hex (#333)'
            }),
            PropertyPaneTextField('AreasExibicao', { 
              label: 'Filtrar por Áreas',
              description: 'Ex: Geral, TI, RH (separe por vírgula)'
            })
          ]
        }]
      }]
    };
  }
}