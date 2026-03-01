import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AniversariosWebPart.module.scss';

export interface IAniversariantesWebPartProps {
  filtroTempo: 'dia' | 'semana' | 'mes';
  corFundo: string;
  imagemFundo: string;
  corTitulo: string;
  corCard: string;
}

interface IAniversarianteItem {
  Id: number;
  Title: string;
  DatadeNascimento: string;
  Foto?: any;
  Email: string;
  Funcao?: string; // Nova propriedade para o cargo
}

export default class AniversariosWebPart extends BaseClientSideWebPart<IAniversariantesWebPartProps> {
  private _sp: SPFI;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
  }

  public render(): void {
    this._obterAniversariantes()
      .then(itens => this._desenharInterface(itens))
      .catch((err) => {
        console.error(err);
        this._desenharInterface([]);
      });
  }

  private _desenharInterface(itens: IAniversarianteItem[]): void {
    const corFundoPainel: string = this.properties.corFundo || "#ffffff";
    const corTextoTitulo: string = this.properties.corTitulo || "#FFFFFF";
    const urlFundo = this.properties.imagemFundo ? `url('${this.properties.imagemFundo}')` : 'none';

    let tituloDinamico = "Aniversariantes do Mês";
    if (this.properties.filtroTempo === 'dia') tituloDinamico = "Aniversariantes do Dia";
    else if (this.properties.filtroTempo === 'semana') tituloDinamico = "Aniversariantes da Semana";

    const conteudo = itens.length > 0
      ? itens.map(item => this._renderItem(item)).join('')
      : `<p style="margin:0; color:${corTextoTitulo}; width:100%; text-align:center; font-weight:bold; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">Nenhum aniversariante encontrado.</p>`;

    this.domElement.innerHTML = `
      <div class="${styles.aniversarios}" style="
        background-color: ${corFundoPainel}; 
        background-image: ${urlFundo};
        background-size: cover;
        background-position: center;
        padding: 25px; 
        border-radius: 15px; 
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
      ">
        <h2 style="
          margin: 0 0 20px 0; 
          color: ${corTextoTitulo}; 
          border-bottom: 2px solid rgba(255,255,255,0.3); 
          padding-bottom: 10px; 
          text-align: center;
          font-weight: bold;
          text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
        ">
          🎂 ${tituloDinamico}
        </h2>
        <div style="display: flex; flex-wrap: wrap; gap: 20px; justify-content: center;">
          ${conteudo}
        </div>
      </div>`;
  }

  private _renderItem(item: IAniversarianteItem): string {
    const defaultFoto = 'https://static.sharepointonline.com/static/beta/sp-client/images/person_default.png';
    const fotoUrl = this._resolverFotoUrl(item.Foto, item.Id) || defaultFoto;
    const corFundoCard: string = this.properties.corCard || "rgba(255,255,255,0.9)";
    
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const bday = this._toLocalDate(item.DatadeNascimento);
    const bdayEsteAno = new Date(hoje.getFullYear(), bday.getMonth(), bday.getDate());
    
    const diffTime = bdayEsteAno.getTime() - hoje.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    const deveMostrarBotao = diffDays >= 0 && diffDays <= 7;

    const mensagem = encodeURIComponent(`Olá ${item.Title}, Feliz Aniversário! 🎂🎉`);
    const teamsUrl = item.Email 
      ? `https://teams.microsoft.com/l/chat/0/0?users=${item.Email}&message=${mensagem}`
      : `#`;

    return `
      <div style="background: ${corFundoCard}; padding: 15px; border-radius: 12px; text-align: center; width: 170px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); position: relative; display: flex; flex-direction: column; align-items: center;">
        <div style="width: 80px; height: 80px; border-radius: 50%; overflow: hidden; margin-bottom: 10px; border: 3px solid #0078d4;">
          <img src="${fotoUrl}" alt="${escape(item.Title)}" style="width: 100%; height: 100%; object-fit: cover;" onerror="this.src='${defaultFoto}';" />
        </div>
        
        <div style="font-size: 14px; font-weight: 600; color: #333; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; width: 100%;" title="${escape(item.Title)}">
          ${escape(item.Title)}
        </div>

        <div style="font-size: 11px; color: #666; margin-top: 2px; font-style: italic; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; width: 100%;" title="${escape(item.Funcao || '')}">
          ${escape(item.Funcao || 'Colaborador')}
        </div>

        <div style="color: #0078d4; font-weight: bold; font-size: 13px; margin-top: 5px;">${this._formatarData(item.DatadeNascimento)}</div>
        
        <div style="min-height: 40px; display: flex; align-items: flex-end;">
          ${deveMostrarBotao && item.Email ? `
            <a href="${teamsUrl}" target="_blank" title="Enviar parabéns no Teams" style="
              text-decoration: none;
              background-color: #0078d4;
              color: white;
              padding: 6px 12px;
              border-radius: 5px;
              font-size: 11px;
              font-weight: bold;
            ">
              💬 Parabéns!
            </a>
          ` : ''} 
        </div>
      </div>`;
  }

  private _formatarData(dataStr: string): string {
    if (!dataStr) return "";
    const d = this._toLocalDate(dataStr);
    const diaNum = d.getDate();
    const dia = diaNum < 10 ? '0' + diaNum : diaNum.toString();
    const mesNum = d.getMonth() + 1;
    const mes = mesNum < 10 ? '0' + mesNum : mesNum.toString();
    return `${dia}/${mes}`;
  }

  private _resolverFotoUrl(foto: any, itemId: number): string | null {
    if (!foto) return null;
    try {
      const obj = typeof foto === 'string' ? JSON.parse(foto) : foto;
      if (obj.serverRelativeUrl) return obj.serverRelativeUrl;
      if (obj.fileName) {
        const serverUrl = window.location.origin;
        const webUrl = this.context.pageContext.web.serverRelativeUrl === '/' ? '' : this.context.pageContext.web.serverRelativeUrl;
        return `${serverUrl}${webUrl}/Lists/Aniversariantes/Attachments/${itemId}/${obj.fileName}`;
      }
    } catch {
      return typeof foto === 'string' ? foto : null;
    }
    return null;
  }

  private _toLocalDate(value: string): Date {
    const d = new Date(value);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  private async _obterAniversariantes(): Promise<IAniversarianteItem[]> {
    try {
      // Adicionado "Funcao" no select
      const lista = await this._sp.web.lists.getByTitle("Aniversariantes").items
        .select("Id", "Title", "DatadeNascimento", "Foto", "Email", "Funcao")();
      
      const hoje = new Date();
      hoje.setHours(0, 0, 0, 0);

      const filtrados = (lista as IAniversarianteItem[]).filter(item => {
        if (!item.DatadeNascimento) return false;
        const bday = this._toLocalDate(item.DatadeNascimento);
        const bdayThisYear = new Date(hoje.getFullYear(), bday.getMonth(), bday.getDate());

        switch (this.properties.filtroTempo) {
          case 'dia': return bday.getDate() === hoje.getDate() && bday.getMonth() === hoje.getMonth();
          case 'semana': {
            const diff = (bdayThisYear.getTime() - hoje.getTime()) / (1000 * 3600 * 24);
            return diff >= 0 && diff <= 7;
          }
          default: return bday.getMonth() === hoje.getMonth();
        }
      });
      return filtrados.sort((a, b) => this._toLocalDate(a.DatadeNascimento).getDate() - this._toLocalDate(b.DatadeNascimento).getDate());
    } catch { return []; }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        groups: [{
          groupFields: [
            PropertyPaneChoiceGroup('filtroTempo', {
              label: 'Período',
              options: [
                { key: 'dia', text: 'Hoje' },
                { key: 'semana', text: 'Próximos 7 dias' },
                { key: 'mes', text: 'Este mês' }
              ]
            }),
            PropertyPaneTextField('corFundo', { label: 'Cor de Fundo do Painel' }),
            PropertyPaneTextField('imagemFundo', { label: 'Link da Imagem de Fundo (URL)' }),
            PropertyPaneTextField('corTitulo', { label: 'Cor do Título' }),
            PropertyPaneTextField('corCard', { label: 'Cor do Card (Ex: rgba(255,255,255,0.8))' })
          ]
        }]
      }]
    };
  }
}