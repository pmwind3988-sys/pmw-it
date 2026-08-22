// The ONLY place ECharts may be imported from in this feature.
//
// Importing the `echarts` umbrella anywhere pulls in every chart type,
// every component and both renderers -- roughly a megabyte, most of it
// for charts this app never draws. That would defeat the whole bundle
// strategy in spec §6.3, and it would do so silently: the code works,
// it is just enormous. Registering an explicit set here is what keeps
// the tree shaking honest.

import * as echarts from 'echarts/core';
import {
  BarChart, LineChart, PieChart, ScatterChart, HeatmapChart,
} from 'echarts/charts';
import {
  GridComponent, TooltipComponent, LegendComponent,
  DataZoomComponent, TitleComponent, VisualMapComponent,
} from 'echarts/components';
import { CanvasRenderer } from 'echarts/renderers';

echarts.use([
  BarChart, LineChart, PieChart, ScatterChart, HeatmapChart,
  GridComponent, TooltipComponent, LegendComponent,
  DataZoomComponent, TitleComponent, VisualMapComponent,
  CanvasRenderer,
]);

export default echarts;
