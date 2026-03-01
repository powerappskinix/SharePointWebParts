import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

export interface IMyCallendarTeamsWebPartProps {
  corTema: string;
}

export default class MyCallendarTeamsWebPart extends BaseClientSideWebPart<IMyCallendarTeamsWebPartProps> {
  private _dataSelecionada: Date = new Date();
  private _eventosAtuais: any[] = [];

  public render(): void {
    const corPrincipal = this.properties.corTema || '#0078d4'; 
    this._renderizarEstruturaBase(corPrincipal);
    this._carregarEventos(corPrincipal).catch(err => console.error(err));
  }

  private _renderizarEstruturaBase(cor: string): void {
    const mesOriginal = this._dataSelecionada.toLocaleDateString('pt-BR', { month: 'long' });
    const anoOriginal = this._dataSelecionada.getFullYear();
    const nomeMesFormatado = `${mesOriginal.charAt(0).toUpperCase() + mesOriginal.slice(1)} ${anoOriginal}`;
    
    this.domElement.innerHTML = `
      <div style="font-family: 'Segoe UI', system-ui; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
          <div style="display: flex; gap: 10px;">
            <button id="btnAnterior" title="Mês Anterior" style="cursor:pointer; width:35px; height:35px; background:#f3f2f1; color:${cor}; border:none; border-radius:50%; font-size:18px; font-weight:bold; display:flex; align-items:center; justify-content:center;">&#10094;</button>
            <button id="btnProximo" title="Próximo Mês" style="cursor:pointer; width:35px; height:35px; background:#f3f2f1; color:${cor}; border:none; border-radius:50%; font-size:18px; font-weight:bold; display:flex; align-items:center; justify-content:center;">&#10095;</button>
          </div>
          <h3 style="margin: 0; font-size: 20px; color: #323130; font-weight: 600;">${escape(nomeMesFormatado)}</h3>
          <div style="width: 80px;"></div>
        </div>
        <div id="calendarioGrid"></div>
        <div id="painelDetalhes" style="margin-top: 20px; padding: 15px; border-top: 4px solid ${cor}; background: #faf9f8; display: none; border-radius: 4px;">
          <div id="listaDetalhes"></div>
        </div>
      </div>`;

    this.domElement.querySelector('#btnAnterior')?.addEventListener('click', () => this._mudarMes(-1));
    this.domElement.querySelector('#btnProximo')?.addEventListener('click', () => this._mudarMes(1));
  }

  private _mudarMes(direcao: number): void {
    this._dataSelecionada.setMonth(this._dataSelecionada.getMonth() + direcao);
    this.render();
  }

  private async _carregarEventos(cor: string): Promise<void> {
    const ano = this._dataSelecionada.getFullYear();
    const mes = this._dataSelecionada.getMonth();
    const inicio = new Date(ano, mes, 1, 0, 0, 0).toISOString();
    const fim = new Date(ano, mes + 1, 0, 23, 59, 59).toISOString();

    try {
      const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

      // Busca calendários para identificar o de Feriados
      const calsResponse = await client.api('/me/calendars').select('id,name').get();
      const calFeriados = calsResponse.value.find((c: any) => 
        c.name.toLowerCase().includes('feriados') || c.name.toLowerCase().includes('holidays')
      );

      // Busca eventos normais
      const resEventos = await client
        .api('/me/calendar/calendarView')
        .header('Prefer', 'outlook.timezone="E. South America Standard Time"')
        .query({ startDateTime: inicio, endDateTime: fim })
        .top(999)
        .get();

      let todosEventos = resEventos.value || [];

      // Busca feriados se o calendário existir
      if (calFeriados) {
        const resFeriados = await client
          .api(`/me/calendars/${calFeriados.id}/calendarView`)
          .query({ startDateTime: inicio, endDateTime: fim })
          .get();
        
        const feriadosMarcados = resFeriados.value.map((f: any) => ({ ...f, isFeriado: true }));
        todosEventos = [...todosEventos, ...feriadosMarcados];
      }
      
      this._eventosAtuais = todosEventos;
      this._desenharGrade(cor);
    } catch (error) {
      console.error("Erro ao carregar dados do Graph:", error);
    }
  }

  private _desenharGrade(cor: string): void {
    const ano = this._dataSelecionada.getFullYear();
    const mes = this._dataSelecionada.getMonth();
    const primeiroDiaMes = new Date(ano, mes, 1);
    const ultimoDiaMes = new Date(ano, mes + 1, 0);
    const diasSemana = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
    
    let html = `<div style="display: grid; grid-template-columns: repeat(7, 1fr); width: 100%; background: #edebe9; border-top: 1px solid #edebe9; border-left: 1px solid #edebe9; box-sizing: border-box;">`;
    
    diasSemana.forEach(dia => {
      html += `<div style="background: #f3f2f1; padding: 10px 0; text-align: center; font-weight: 600; font-size: 12px; color: #605e5c; border-right: 1px solid #edebe9; border-bottom: 1px solid #edebe9;">${dia}</div>`;
    });

    for (let i = 0; i < primeiroDiaMes.getDay(); i++) {
      html += '<div style="background: #ffffff; min-height: 100px; border-right: 1px solid #edebe9; border-bottom: 1px solid #edebe9;"></div>';
    }

    for (let dia = 1; dia <= ultimoDiaMes.getDate(); dia++) {
      const dataFocoStr = new Date(ano, mes, dia).toLocaleDateString('pt-BR');
      const evs = this._eventosAtuais.filter(e => new Date(e.start.dateTime).toLocaleDateString('pt-BR') === dataFocoStr);
      const eHoje = new Date().toLocaleDateString('pt-BR') === dataFocoStr;

      html += `
        <div class="dia-calendario" data-data="${dataFocoStr}" style="background: ${eHoje ? '#f4f4fc' : '#fff'}; min-height: 110px; padding: 5px; border-right: 1px solid #edebe9; border-bottom: 1px solid #edebe9; cursor: pointer; box-sizing: border-box; overflow: hidden;">
          <div style="font-size: 12px; font-weight: ${eHoje ? '700' : '400'}; color: ${eHoje ? cor : '#323130'};">${dia}</div>
          <div style="margin-top: 5px; display: flex; flex-direction: column; gap: 2px;">
            ${evs.slice(0, 3).map(e => {
              const hora = new Date(e.start.dateTime).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
              const corItem = e.isFeriado ? '#8a8886' : cor; 
              const textoExibir = e.isFeriado ? escape(e.subject) : `${hora} ${escape(e.subject)}`;
              return `<div style="background: ${corItem}; color: #fff; font-size: 9px; padding: 1px 3px; border-radius: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${textoExibir}</div>`;
            }).join('')}
            ${evs.length > 3 ? `<div style="font-size: 9px; color: ${cor}; font-weight: bold;">+ ${evs.length - 3} mais</div>` : ''}
          </div>
        </div>`;
    }

    const gridElem = this.domElement.querySelector('#calendarioGrid');
    if (gridElem) {
      gridElem.innerHTML = html + '</div>';
      this.domElement.querySelectorAll('.dia-calendario').forEach(el => {
        el.addEventListener('click', () => this._exibirDetalhes(el.getAttribute('data-data') || "", cor));
      });
    }
  }

  private _exibirDetalhes(data: string, cor: string): void {
    const painel = this.domElement.querySelector('#painelDetalhes') as HTMLElement;
    const lista = this.domElement.querySelector('#listaDetalhes');
    const filtrados = this._eventosAtuais.filter(e => new Date(e.start.dateTime).toLocaleDateString('pt-BR') === data);

    if (painel && lista) {
      painel.style.display = 'block';
      lista.innerHTML = `<h4 style="margin-top: 0; color: ${cor};">Eventos de ${escape(data)}</h4>`;
      
      if (filtrados.length === 0) {
        lista.innerHTML += `<p style="font-size: 14px; color: #605e5c;">Nenhuma reunião agendada.</p>`;
      } else {
        filtrados.forEach(e => {
          const inicio = new Date(e.start.dateTime);
          const fim = new Date(e.end.dateTime);
          const duracao = Math.round((fim.getTime() - inicio.getTime()) / 60000);
          const corTexto = e.isFeriado ? '#8a8886' : '#323130';

          lista.innerHTML += `
            <div style="padding: 10px; border-bottom: 1px solid #edebe9; margin-bottom: 5px; background: #fff; border-radius: 4px;">
              <strong style="color: ${corTexto}; font-size: 14px;">${escape(e.subject)} ${e.isFeriado ? '(Feriado)' : ''}</strong><br/>
              ${!e.isFeriado ? `
                <span style="font-size: 12px; color: #605e5c;">
                  ⏰ ${inicio.toLocaleTimeString('pt-BR', {hour:'2-digit', minute:'2-digit'})} - ${fim.toLocaleTimeString('pt-BR', {hour:'2-digit', minute:'2-digit'})} 
                  <br/>⏳ Duração: <strong>${duracao} minutos</strong>
                </span>
              ` : ''}
            </div>`;
        });
      }
      painel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Configure seu Calendário" },
          groups: [
            {
              groupName: "Cores",
              groupFields: [
                PropertyPaneTextField('corTema', {
                  label: 'Código Hex da Cor (Ex: #ff0000)',
                  description: 'Defina a cor principal da identidade visual'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }
}