import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle // Adicionado o Toggle
} from '@microsoft/sp-property-pane';


export interface IInstagramFeedWebPartProps {
  instaAccount: string;
  alturaFeed: number;
  larguraFeed: number; 
  larguraTotal: boolean;
}

export default class InstagramFeedWebPart extends BaseClientSideWebPart<IInstagramFeedWebPartProps> {

  public render(): void {
    const user = this.properties.instaAccount || "microsoft";
    const altura = this.properties.alturaFeed || 600;
    const largura = this.properties.larguraFeed || 400; 
    const larguraTotal = this.properties.larguraTotal;

    // Remove paddings padrão do container do SharePoint
    this.domElement.style.margin = '0';
    this.domElement.style.padding = '0';
    this.domElement.style.width = '100%';

    // Define se usa largura fixa ou 100%
    const larguraFinal = larguraTotal ? '100%' : `${largura}px`;

    this.domElement.innerHTML = `
      <div style="
        width: ${larguraFinal}; 
        margin: 0; 
        padding: 15px; 
        background: #fff; 
        border-radius: 8px; 
        box-shadow: 0 4px 12px rgba(0,0,0,0.1); 
        font-family: 'Segoe UI', system-ui;
        box-sizing: border-box;
      ">
        
        <div style="width: 100%; background: #faf9f8; overflow: hidden;">
          <iframe 
            src="https://www.instagram.com/${user}/embed" 
            width="100%" 
            height="${altura}" 
            frameborder="0" 
            scrolling="no" 
            allowtransparency="true"
            style="border: 1px solid #edebe9;">
          </iframe>
        </div>
      </div>`;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Ajuste de Dimensões" },
          groups: [
            {
              groupName: "Configurações Gerais",
              groupFields: [
                PropertyPaneTextField('instaAccount', {
                  label: 'Username do Instagram'
                }),
                PropertyPaneToggle('larguraTotal', {
                  label: 'Ocupar largura total da coluna',
                  onText: 'Sim (100%)',
                  offText: 'Não (Usar Slider)'
                }),
                PropertyPaneSlider('larguraFeed', {
                  label: 'Largura da Web Part (px)',
                  min: 300,
                  max: 800,
                  step: 10,
                  disabled: this.properties.larguraTotal // Desativa o slider se largura total estiver ativa
                }),
                PropertyPaneSlider('alturaFeed', {
                  label: 'Altura do Feed (px)',
                  min: 300,
                  max: 1000,
                  step: 50
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}