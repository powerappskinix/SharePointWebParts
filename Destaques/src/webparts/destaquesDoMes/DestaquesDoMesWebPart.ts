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

import styles from './DestaquesDoMesWebPart.module.scss';

export interface IDestaquesDoMesWebPartProps {
  corFundo: string;
  imagemFundo: string;
  corTitulo: string;
  corCard: string;
  // Títulos
  titAdmissao: string;
  titPromocao: string;
  titMovimentacao: string;
  titDesligamento: string;
  // Subtítulos
  subAdmissao: string;
  subPromocao: string;
  subMovimentacao: string;
  subDesligamento: string;
  // Cores
  corPromocao: string;
  corAdmissao: string;
  corMovimentacao: string;
  corDesligamento: string;
}

interface IDestaqueItem {
  Id: number;
  Title: string;
  Foto?: any;
  Funcao: string;
  Departamento: string;
  TipoDestaque: string;
  Email: string;
}

export default class DestaquesDoMesWebPart extends BaseClientSideWebPart<IDestaquesDoMesWebPartProps> {
  private _sp: SPFI;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
  }

  public render(): void {
    this._obterDestaques()
      .then(itens => this._desenharInterface(itens))
      .catch((err) => {
        console.error("Erro no render:", err);
        this._desenharInterface([]);
      });
  }

  private _desenharInterface(itens: IDestaqueItem[]): void {
    const corFundoPainel = this.properties.corFundo || "#ffffff";
    const urlFundo = this.properties.imagemFundo ? `url('${this.properties.imagemFundo}')` : 'none';

    const admissoes = itens.filter(i => this._checkTipo(i.TipoDestaque, 'admiss'));
    const promocoes = itens.filter(i => this._checkTipo(i.TipoDestaque, 'promo'));
    const movimentacoes = itens.filter(i => this._checkTipo(i.TipoDestaque, 'moviment'));
    const desligamentos = itens.filter(i => this._checkTipo(i.TipoDestaque, 'desliga'));

    this.domElement.innerHTML = `
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Asap:wght@400;700&display=swap');
        .destaque-asap-font { font-family: 'Asap', sans-serif !important; }
      </style>
      <div class="${styles.destaques} destaque-asap-font" style="
        background-color: ${corFundoPainel}; 
        background-image: ${urlFundo};
        background-size: cover;
        background-position: center;
        padding: 30px 15px; 
        border-radius: 15px;
      ">
        <div style="display: flex; flex-direction: column; gap: 40px;">
          ${this._renderSecao(
            this.properties.titAdmissao || "Novos Integrantes 🤝", 
            this.properties.subAdmissao || "", 
            admissoes, 
            this.properties.corAdmissao || "#0078d4"
          )}
          ${this._renderSecao(
            this.properties.titPromocao || "Promoções 🚀", 
            this.properties.subPromocao || "", 
            promocoes, 
            this.properties.corPromocao || "#28a745"
          )}
          ${this._renderSecao(
            this.properties.titMovimentacao || "Movimentações 🔄", 
            this.properties.subMovimentacao || "", 
            movimentacoes, 
            this.properties.corMovimentacao || "#FF8C00"
          )}
          ${this._renderSecao(
            this.properties.titDesligamento || "Ciclos que se Encerram 👋", 
            this.properties.subDesligamento || "", 
            desligamentos, 
            this.properties.corDesligamento || "#d11d12"
          )}
        </div>
      </div>`;
  }

  private _renderSecao(titulo: string, subtitulo: string, itens: IDestaqueItem[], corDestaque: string): string {
    if (itens.length === 0) return "";
    return `
      <div>
        <div style="margin: 0 0 15px 10px; border-left: 5px solid ${corDestaque}; padding-left: 12px;">
            <h3 style="color: ${this.properties.corTitulo || '#fff'}; margin: 0; text-transform: uppercase; font-size: 18px; font-weight: 700;">${escape(titulo)}</h3>
            ${subtitulo ? `<p style="color: ${this.properties.corTitulo || '#fff'}; margin: 2px 0 0 0; font-size: 13px; font-weight: 400; opacity: 0.9;">${escape(subtitulo)}</p>` : ''}
        </div>
        <div style="display: flex; flex-wrap: wrap; gap: 15px; justify-content: flex-start;">
          ${itens.map(item => this._renderItem(item)).join('')}
        </div>
      </div>`;
  }

  private _renderItem(item: IDestaqueItem): string {
    const defaultFoto = 'https://static.sharepointonline.com/static/beta/sp-client/images/person_default.png';
    const fotoUrl = this._resolverFotoUrl(item.Foto, item.Id) || defaultFoto;
    const corFundoCard = this.properties.corCard || "rgba(255,255,255,0.95)";
    const tipo = (item.TipoDestaque || "").toLowerCase();
    
    let seloCor = this.properties.corAdmissao || "#0078d4";
    let textoBotao = "Boas-vindas!";
    let msgTexto = 'É um prazer ter você conosco, ${item.Title}! 🎉. Desejamos que sua jornada aqui seja de muito aprendizado, crescimento e ótimas conquistas. Conte com nossa equipe sempre que precisar. ';

    if (tipo.indexOf('promo') !== -1) {
      seloCor = this.properties.corPromocao || "#28a745";
      textoBotao = "Parabéns!";
      msgTexto = `Seu esforço e dedicação foram reconhecidos, ${item.Title}! 🎉.Desejamos muito sucesso nesta nova fase, você merece! `;
    } else if (tipo.indexOf('moviment') !== -1) {
      seloCor = this.properties.corMovimentacao || "#FF5C00";
      textoBotao = "Boa sorte!";
      msgTexto = `Agradecemos o trabalho realizado e desejamos sucesso na nova função , ${item.Title}! `;
    } else if (tipo.indexOf('desliga') !== -1) {
      seloCor = this.properties.corDesligamento || "#3C3C3C";
      textoBotao = "Até breve!";
      msgTexto = `Agradecemos o período em que fez parte da equipe e desejamos sucesso em seus próximos caminhos, ${item.Title}!`;
    }

    const teamsUrl = item.Email ? `https://teams.microsoft.com/l/chat/0/0?users=${item.Email}&message=${encodeURIComponent(msgTexto)}` : `#`;

    return `
      <div style="background: ${corFundoCard}; padding: 15px 10px; border-radius: 10px; text-align: center; width: 175px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); border-top: 5px solid ${seloCor};">
        <div style="width: 75px; height: 75px; border-radius: 50%; overflow: hidden; margin: 8px auto; border: 2px solid ${seloCor};">
          <img src="${fotoUrl}" style="width: 100%; height: 100%; object-fit: cover;" onerror="this.src='${defaultFoto}';" />
        </div>
        <div style="font-size: 14px; font-weight: 700; color: #333; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${escape(item.Title)}</div>
        <div style="font-size: 11px; color: #555; font-weight: 400; margin-top: 2px;">${escape(item.Funcao || '')}</div>
        <div style="font-size: 10px; color: #777; font-weight: 400; font-style: italic;">${escape(item.Departamento || 'Geral')}</div>
        <div style="margin-top: 12px;">
          <a href="${teamsUrl}" target="_blank" style="text-decoration: none; background-color: ${seloCor}; color: white; padding: 6px 12px; border-radius: 4px; font-size: 10px; font-weight: 700; display: inline-block; width: 85%;">
            ${textoBotao}
          </a>
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
        return `${serverUrl}${webUrl}/Lists/Movimentao%20de%20Pessoas/Attachments/${itemId}/${obj.fileName}`;
      }
    } catch { return typeof foto === 'string' ? foto : null; }
    return null;
  }

  private _checkTipo(valor: string, busca: string): boolean {
    if (!valor) return false;
    return valor.toLowerCase().indexOf(busca) !== -1;
  }

  private async _obterDestaques(): Promise<IDestaqueItem[]> {
    try {
      return await this._sp.web.lists.getByTitle("Movimentação de Pessoas").items
        .select("Id", "Title", "Foto", "Funcao", "Departamento", "TipoDestaque", "Email")();
    } catch (error) {
      console.error("Erro na lista:", error);
      return [];
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        groups: [
          {
            groupName: "Configurações Gerais",
            groupFields: [
              PropertyPaneTextField('corFundo', { label: 'Cor de Fundo' }),
              PropertyPaneTextField('imagemFundo', { label: 'URL Imagem Fundo' }),
              PropertyPaneTextField('corTitulo', { label: 'Cor dos Textos da Seção' }),
              PropertyPaneTextField('corCard', { label: 'Cor do Card' })
            ]
          },
          {
            groupName: "Seção: Admissão",
            groupFields: [
              PropertyPaneTextField('titAdmissao', { label: 'Título' }),
              PropertyPaneTextField('subAdmissao', { label: 'Subtítulo' }),
              PropertyPaneTextField('corAdmissao', { label: 'Cor' })
            ]
          },
          {
            groupName: "Seção: Promoção",
            groupFields: [
              PropertyPaneTextField('titPromocao', { label: 'Título' }),
              PropertyPaneTextField('subPromocao', { label: 'Subtítulo' }),
              PropertyPaneTextField('corPromocao', { label: 'Cor' })
            ]
          },
          {
            groupName: "Seção: Movimentação",
            groupFields: [
              PropertyPaneTextField('titMovimentacao', { label: 'Título' }),
              PropertyPaneTextField('subMovimentacao', { label: 'Subtítulo' }),
              PropertyPaneTextField('corMovimentacao', { label: 'Cor' })
            ]
          },
          {
            groupName: "Seção: Desligamento",
            groupFields: [
              PropertyPaneTextField('titDesligamento', { label: 'Título' }),
              PropertyPaneTextField('subDesligamento', { label: 'Subtítulo' }),
              PropertyPaneTextField('corDesligamento', { label: 'Cor' })
            ]
          }
        ]
      }]
    };
  }
}