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

import styles from './DestaquesEmpresaWebPart.module.scss';

export interface IDestaquesEmpresaProps {
  tituloWebPart: string;
  corTexto: string;
  corFundo: string;
}

interface IGaleriasItem {
  Id: number;
  Title: string;
  Area: string;
  Conquista: string;
  Foto?: any;
}

export default class DestaquesEmpresaWebPart extends BaseClientSideWebPart<IDestaquesEmpresaProps> {
  private _sp: SPFI;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
  }

  public render(): void {
    this._obterItens()
      .then(itens => {
        const corTexto = this.properties.corTexto || "#333333";
        const corFundo = this.properties.corFundo || "#ffffff";
        const tituloWS = this.properties.tituloWebPart || "Destaques";

        const listaHtml = itens.length > 0 
          ? itens.map(item => this._renderLinha(item, corTexto)).join('')
          : `<p style="color:${corTexto}; font-size:17px; font-family: 'Asap', sans-serif;">Nenhum destaque para exibir.</p>`;

        this.domElement.innerHTML = `
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Asap:ital,wght@0,400;0,700;1,400;1,700&display=swap');
        
        /* Classe para garantir que os itens tenham uma largura mínima e cresçam conforme o espaço */
        .item-destaque {
          display: flex; 
          align-items: center; 
          padding: 10px; 
          min-width: 300px; /* Largura mínima para não ficar muito espremido */
          flex: 1 1 300px;  /* Permite que o item cresça e ocupe o espaço disponível */
          border: none; 
          box-shadow: none; 
          font-family: 'Asap', sans-serif;
        }
      </style>

      <div class="${styles.destaquesEmpresa}" style="padding: 30px; background-color: ${corFundo}; border: 1px solid #e1e1e1; border-radius: 4px; box-shadow: none; font-family: 'Asap', sans-serif;">
        
        <div style="display: flex; align-items: center; margin-bottom: 30px;">
          <span style="font-size: 35px; margin-right: 15px;">⭐</span>
          <h1 style="font-size: 28px; font-weight: 700; margin: 0; font-family: 'Asap', sans-serif; color: ${corTexto}; border: none; box-shadow: none;">
            ${escape(tituloWS)}
          </h1>
        </div>

        <div style="display: flex; flex-direction: row; flex-wrap: wrap; gap: 25px;">
          ${listaHtml}
        </div>

      </div>`;
          });
  }

  private _renderLinha(item: IGaleriasItem, cor: string): string {
    const fotoUrl = this._resolverFotoUrl(item.Foto, item.Id) || 'https://static.sharepointonline.com/static/beta/sp-client/images/person_default.png';
    const tamanhoFonte = "17px"; 

    return `
      <div style="display: flex; align-items: center; padding: 0; margin: 0; border: none; box-shadow: none; font-family: 'Asap', sans-serif;">
        <div style="width: 50px; height: 50px; border-radius: 50%; overflow: hidden; margin-right: 20px; flex-shrink: 0; border: 1px solid rgba(0,0,0,0.05);">
          <img src="${fotoUrl}" style="width: 100%; height: 100%; object-fit: cover;" />
        </div>
        
        <div style="color: ${cor}; border: none; box-shadow: none;">
          <div style="font-size: ${tamanhoFonte}; font-weight: 400; margin-bottom: 2px;">
            ${escape(item.Area)} - ${escape(item.Conquista)}
          </div>
          <div style="font-size: ${tamanhoFonte}; font-weight: 400;">
            ${escape(item.Title)}
          </div>
        </div>
      </div>`;
  }

  private _resolverFotoUrl(foto: any, itemId: number): string | null {
    if (!foto) return null;
    try {
      const obj = typeof foto === 'string' ? JSON.parse(foto) : foto;
      if (obj.serverRelativeUrl) return obj.serverRelativeUrl;
      if (obj.fileName) {
        const serverUrl = window.location.origin;
        const webUrl = this.context.pageContext.web.serverRelativeUrl === '/' ? '' : this.context.pageContext.web.serverRelativeUrl;
        return `${serverUrl}${webUrl}/Lists/Destaques/Attachments/${itemId}/${obj.fileName}`;
      }
    } catch { return typeof foto === 'string' ? foto : null; }
    return null;
  }

  private async _obterItens(): Promise<IGaleriasItem[]> {
    try {
      return await this._sp.web.lists.getByTitle("Destaques").items
        .select("Id", "Title", "Area", "Conquista", "Foto")();
    } catch { return []; }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        groups: [{
          groupName: "Configurações de Aparência",
          groupFields: [
            PropertyPaneTextField('tituloWebPart', { label: 'Título Principal' }),
            PropertyPaneTextField('corTexto', { label: 'Cor do Texto' }),
            PropertyPaneTextField('corFundo', { label: 'Cor de Fundo' })
          ]
        }]
      }]
    };
  }
}