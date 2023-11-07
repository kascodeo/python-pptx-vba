from ..enum.XlAxisGroup import XlAxisGroup


class Chart():
    def __init__(self, parent, chartspace):
        self._parent = parent
        self._chartspace = chartspace
        self._chartarea = None
        self._chartdata = None
        self._datatable = None
        self._chartformat = None
        self._plotarea = None
        self._legend = None
        self._catax = None
        self._valax = None
        self._serax = None

    def Axes(self, Type=None, AxisGroup=XlAxisGroup.xlPrimary):
        pass

    def SetSourceData(self):
        pass

    @property
    def ChartArea(self):
        from .chartarea import ChartArea
        if self._chartarea is None:
            self._chartarea = ChartArea(self)
        return self._chartarea

    @property
    def ChartColor(self):
        pass

    @property
    def ChartData(self):
        from .chartdata import ChartData
        if self._chartdata is None:
            self._chartdata = ChartData(self)
        return self._chartdata

    @property
    def ChartStyle(self):
        pass

    @property
    def ChartTitle(self):
        from .charttitle import ChartTitle
        if self._charttitle is None:
            self._charttitle = ChartTitle(self)
        return self._charttitle

    @property
    def ChartType(self):
        pass

    @property
    def DataTable(self):
        from .datatable import DataTable
        if self._datatable is None:
            self._datatable = DataTable(self)
        return self._datatable

    @property
    def Format(self):
        from .chartformat import ChartFormat
        if self._chartformat is None:
            self._chartformat = ChartFormat(self, None)
        return self._chartformat

    @property
    def GapDepth(self):
        pass

    @property
    def HasAxis(self):
        return any(map(lambda i: i is not None, [self.e_catAx, self.e_valAx]))

    @property
    def HasDataTable(self):
        return self.e_dTable is not None

    @property
    def HasLegend(self):
        return self.e_legend is not None

    @property
    def HasTitle(self):
        return self.e_title is not None

    @property
    def Legend(self):
        from .legend import Legend
        if self._legend is None:
            self._legend = Legend(self, self.e_legend)
        return self._legend

    @property
    def Name(self):
        pass

    @property
    def Parent(self):
        return self._parent

    @property
    def PlotArea(self):
        from .plotarea import PlotArea
        if self._plotarea is None:
            self._plotarea = PlotArea(self, self.e_plotArea)
        return self._plotarea

    @property
    def RightAngleAxes(self):
        pass

    @property
    def SeriesNameLevel(self):
        pass

    @property
    def Shapes(self):
        pass

    @property
    def Title(self):  # string return
        pass

    @property
    def e_chartspace(self):
        return self._chartspace.e

    @property
    def e_chart(self):
        return self.e_chartspace.findqn('c:chart')

    @property
    def e_title(self):
        return self.e_chart.findqn('c:title')

    @property
    def e_plotArea(self):
        return self.e_chart.findqn('c:plotArea')

    @property
    def e_barChart(self):
        return self.e_plotArea.findqn('c:barChart')

    @property
    def e_catAx(self):
        return self.e_plotArea.findqn('c:catAx')

    @property
    def e_valAx(self):
        return self.e_plotArea.findqn('c:valAx')

    @property
    def e_dTable(self):
        return self.e_plotArea.findqn('c:dTable')

    @property
    def e_legend(self):
        return self.e_chart.findqn('c:legend')
