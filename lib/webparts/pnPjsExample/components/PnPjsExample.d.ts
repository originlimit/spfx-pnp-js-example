import * as React from 'react';
import { IPnPjsExampleProps } from './IPnPjsExampleProps';
import { IFile } from "./interfaces";
export interface IAsyncAwaitPnPJsProps {
    description: string;
}
export interface IIPnPjsExampleState {
    items: IFile[];
    errors: string[];
}
export default class PnPjsExample extends React.Component<IPnPjsExampleProps, IIPnPjsExampleState> {
    private LOG_SOURCE;
    private LIBRARY_NAME;
    private _sp;
    constructor(props: IPnPjsExampleProps);
    componentDidMount(): void;
    render(): React.ReactElement<IAsyncAwaitPnPJsProps>;
    private _readAllFilesSize;
    private _updateTitles;
}
//# sourceMappingURL=PnPjsExample.d.ts.map