import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
export interface IPnPjsExampleWebPartProps {
    description: string;
}
export default class PnPjsExampleWebPart extends BaseClientSideWebPart<IPnPjsExampleWebPartProps> {
    private _isDarkTheme;
    private _environmentMessage;
    render(): void;
    onInit(): Promise<void>;
    private _getEnvironmentMessage;
    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=PnPjsExampleWebPart.d.ts.map