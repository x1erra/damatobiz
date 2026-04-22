declare module 'react-plotly.js' {
    import { Component } from 'react';
    import { PlotData, Layout, Config } from 'plotly.js';

    interface PlotlyProps {
        data: Partial<PlotData>[];
        layout?: Partial<Layout>;
        config?: Partial<Config>;
        style?: React.CSSProperties;
        onInitialized?: (figure: any, graphDiv: HTMLElement) => void;
        onUpdate?: (figure: any, graphDiv: HTMLElement) => void;
        onPurge?: (figure: any, graphDiv: HTMLElement) => void;
        onError?: (err: any) => void;
        useResizeHandler?: boolean;
        className?: string;
    }

    export default class Plot extends Component<PlotlyProps> { }
}
