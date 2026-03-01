import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
export interface IMyCallendarTeamsWebPartProps {
    corTema: string;
}
export default class MyCallendarTeamsWebPart extends BaseClientSideWebPart<IMyCallendarTeamsWebPartProps> {
    private _dataSelecionada;
    private _eventosAtuais;
    render(): void;
    private _renderizarEstruturaBase;
    private _mudarMes;
    private _carregarEventos;
    private _desenharGrade;
    private _exibirDetalhes;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
    protected get dataVersion(): Version;
}
//# sourceMappingURL=MyCallendarTeamsWebPart.d.ts.map