import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';

import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { escape } from '@microsoft/sp-lodash-subset';

export interface IFaqCorporativoWebPartProps {
  listaCategorias: string;
  listaPerguntas: string;
  corPrincipal: string;
  corTitulo: string;
  corFundo: string;
  corDestaqueMenu: string; // Nova propriedade para a cor do Recursos Humanos selecionado
}

export default class FaqCorporativoWebPart extends BaseClientSideWebPart<IFaqCorporativoWebPartProps> {
  private _sp: SPFI;
  private _opcoesListas: IPropertyPaneDropdownOption[] = [];
  private _categorias: any[] = [];
  private _perguntas: any[] = [];

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    if (this._opcoesListas.length > 0) return;
    try {
      const listas = await this._sp.web.lists.filter("Hidden eq false").select("Title")();
      this._opcoesListas = listas.map(l => ({ key: l.Title, text: l.Title }));
      this.context.propertyPane.refresh();
    } catch (e) { console.error(e); }
  }

  public render(): void {
    if (!this.properties.listaCategorias || !this.properties.listaPerguntas) {
      this.domElement.innerHTML = `
        <div style="padding:40px; border:2px dashed #ccc; text-align:center; font-family:'Asap', sans-serif;">
          <h3>Configuração Necessária</h3>
          <p>Selecione as listas de Categorias e Perguntas no painel lateral.</p>
        </div>`;
      return;
    }

    this._obterDados()
      .then(({ categorias, perguntas }) => {
        this._categorias = categorias;
        this._perguntas = perguntas;
        this._renderizarLayoutBase();
      })
      .catch(err => {
        console.error("Erro ao carregar dados:", err);
        this.domElement.innerHTML = `<div style="color:red; padding:20px;">Erro ao carregar dados. Verifique o console (F12).</div>`;
      });
  }

  private async _obterDados(): Promise<{ categorias: any[], perguntas: any[] }> {
    try {
      const catList = this._sp.web.lists.getByTitle(this.properties.listaCategorias);
      const perList = this._sp.web.lists.getByTitle(this.properties.listaPerguntas);

      const categoriasRaw = await catList.items();
      const categorias = categoriasRaw
        .filter(c => c.Visibilidade !== false)
        .sort((a, b) => (Number(a.Sequencia) || 0) - (Number(b.Sequencia) || 0));

      const perguntasRaw = await perList.items();
      const perguntas = perguntasRaw
        .filter(p => p.Visibilidade !== false)
        .sort((a, b) => (Number(a.Sequencia) || 0) - (Number(b.Sequencia) || 0))
        .map(p => ({
          Title: p.Pergunta || p.Title || "Sem título", // Mapeado para Pergunta conforme sua imagem
          Resposta: p.Resposta || "",
          CategoriaId: p.CategoriaId || 0
        }));

      return { categorias, perguntas };
    } catch (error) {
      console.error("Erro no método _obterDados:", error);
      throw error;
    }
  }

  private _renderizarLayoutBase(): void {
    const cor = this.properties.corPrincipal || "#00415a";
    const corT = this.properties.corTitulo || "#333";
    const corF = this.properties.corFundo || "#ffffff";
    const corDM = this.properties.corDestaqueMenu || "#f0f4f7"; 
    
    let html = `
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Asap:wght@400;700&display=swap');
        
        .faq-app { 
          font-family: 'Asap', sans-serif; 
          display: flex; 
          gap: 20px; 
          background: ${corF}; 
          padding: 20px; 
          min-height: 450px;
          border-radius: 8px;
          box-sizing: border-box; 
        }
        
        .faq-sidebar { width: 250px; border-right: 1px solid #eee; padding-right: 15px; flex-shrink: 0; }
        .faq-side-title { font-weight: 700; font-size: 18px; color: ${corT}; margin-bottom: 20px; padding: 10px 0; border-bottom: 2px solid ${cor}; text-transform: uppercase; }
        
        .faq-cat-btn { padding: 12px 15px; cursor: pointer; border-radius: 6px; color: #666; transition: 0.2s; border-left: 4px solid transparent; margin-bottom: 5px; }
        .faq-cat-btn:hover { background: rgba(0,0,0,0.05); }
        
        .faq-cat-btn.active { 
          background: ${corDM}; 
          border-left-color: ${cor}; 
          color: ${cor}; 
          font-weight: 700; 
        }

        .faq-main { flex-grow: 1; padding-left: 10px; overflow: hidden; }
        .faq-search-container { margin-bottom: 25px; width: 100%; box-sizing: border-box; }
        .faq-search-input { 
          width: 100%; 
          padding: 12px 20px; 
          border-radius: 30px; 
          border: 1px solid #ddd; 
          font-family: 'Asap';
          box-sizing: border-box; 
          outline: none;
        }
        .faq-search-input:focus { border-color: ${cor}; }
        
        /* AJUSTE DOS CARDS COM A COR DE DESTAQUE */
        .faq-item { 
          border: 1px solid #e1e1e1; /* Borda padrão leve */
          border-radius: 8px;
          width: 100%; 
          background: #fff; 
          margin-bottom: 10px;
          transition: all 0.3s ease;
          box-sizing: border-box;
        }

        .faq-item:hover {
          border-color: ${cor}; /* Borda assume a cor de destaque no hover */
        }

        .faq-item.is-open { 
          border: 1px solid ${cor}; /* Borda assume a cor de destaque quando aberto */
          box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        }

        .faq-header { 
          padding: 15px 20px; 
          cursor: pointer; 
          font-weight: 700; 
          display: flex; 
          justify-content: space-between; 
          align-items: center; 
          font-size: 16px; 
          color: #333; 
        }

        .faq-chevron { transition: transform 0.3s ease; font-size: 12px; color: ${cor}; font-weight: bold; }
        .faq-body { max-height: 0; overflow: hidden; transition: all 0.3s ease; color: #555; line-height: 1.6; padding: 0 20px; }
        
        .faq-item.is-open .faq-body { 
          max-height: 1500px; 
          padding: 15px 20px; 
          border-top: 1px solid #f5f5f5; 
        }
        .faq-item.is-open .faq-header { color: ${cor}; }
        .faq-item.is-open .faq-chevron { transform: rotate(180deg); }
      </style>

      <div class="faq-app">
        <aside class="faq-sidebar">
          <div class="faq-side-title">Categorias</div>
          <div class="faq-cat-list" id="cat-nav-container">
            ${this._categorias.map((c, i) => `
              <div class="faq-cat-btn ${i === 0 ? 'active' : ''}" data-catid="${c.Id}">
                ${escape(c.Title)}
              </div>
            `).join('')}
          </div>
        </aside>

        <main class="faq-main">
          <div class="faq-search-container">
            <input type="text" class="faq-search-input" id="faq-search-field" placeholder="O que você está procurando?">
          </div>
          <div id="faq-questions-render"></div>
        </main>
      </div>`;

    this.domElement.innerHTML = html;
    this._vincularEventos();

    if (this._categorias.length > 0) {
      this._mostrarPerguntas(this._categorias[0].Id);
    }
  }

  private _vincularEventos(): void {
    this.domElement.querySelectorAll('.faq-cat-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        this.domElement.querySelectorAll('.faq-cat-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const id = parseInt(btn.getAttribute('data-catid') || '0');
        this._mostrarPerguntas(id);
      });
    });

    const search = this.domElement.querySelector('#faq-search-field') as HTMLInputElement;
    if (search) {
      search.addEventListener('input', () => {
        this._mostrarPerguntas(0, search.value.toLowerCase());
      });
    }
  }

  private _mostrarPerguntas(catId: number, busca: string = ""): void {
    const container = this.domElement.querySelector('#faq-questions-render') as HTMLElement;
    if (!container) return;

    let filtradas = this._perguntas;

    if (busca) {
      filtradas = filtradas.filter(p => p.Title.toLowerCase().indexOf(busca) !== -1);
    } else {
      filtradas = filtradas.filter(p => p.CategoriaId === catId);
    }

    container.innerHTML = filtradas.length > 0 ? filtradas.map(p => `
      <div class="faq-item">
        <div class="faq-header">
          <span>${escape(p.Title)}</span>
          <span class="faq-chevron">▼</span>
        </div>
        <div class="faq-body">${p.Resposta}</div>
      </div>
    `).join('') : '<p style="padding:20px; color:#999;">Nenhuma pergunta encontrada.</p>';

    container.querySelectorAll('.faq-header').forEach(header => {
      header.addEventListener('click', () => {
        header.parentElement?.classList.toggle('is-open');
      });
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        groups: [{
          groupName: "Configurações",
          groupFields: [
            PropertyPaneDropdown('listaCategorias', { label: 'Lista de Categorias', options: this._opcoesListas }),
            PropertyPaneDropdown('listaPerguntas', { label: 'Lista de Perguntas', options: this._opcoesListas }),
            PropertyPaneTextField('corPrincipal', { label: 'Cor de Destaque/Linha lateral (HEX)' }),
            PropertyPaneTextField('corTitulo', { label: 'Cor dos Títulos (HEX)' }),
            PropertyPaneTextField('corFundo', { label: 'Cor de Fundo (HEX)' }),
            PropertyPaneTextField('corDestaqueMenu', { label: 'Cor do Fundo da Categoria Selecionada (HEX)' })
          ]
        }]
      }]
    };
  }
}