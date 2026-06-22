export function makeEstateBreakdownSpec(values = {}) {
  const managedAssets = Number(values.managedAssets || 0);
  const realEstate = Number(values.realEstateValue || values.realEstate || 0);
  const otherAssets = Number(values.otherAssets || 0);

  return {
    $schema: "https://vega.github.io/schema/vega-lite/v5.json",
    width: 760,
    height: 260,
    background: "#F7F5F0",
    title: {
      text: "Estate / Net Worth Breakdown",
      anchor: "start",
      fontSize: 22,
      fontWeight: "bold",
      color: "#1A2744"
    },
    data: {
      values: [
        { category: "Managed Assets", value: managedAssets },
        { category: "Real Estate", value: realEstate },
        { category: "Other Assets", value: otherAssets }
      ]
    },
    mark: {
      type: "arc",
      innerRadius: 70,
      outerRadius: 120,
      stroke: "#F7F5F0",
      strokeWidth: 3
    },
    encoding: {
      theta: { field: "value", type: "quantitative" },
      color: {
        field: "category",
        type: "nominal",
        scale: {
          range: ["#3A6BBF", "#B8892A", "#1E7A6E"]
        },
        legend: {
          orient: "right",
          title: null,
          labelFontSize: 15,
          labelColor: "#1A2030"
        }
      },
      tooltip: [
        { field: "category", type: "nominal" },
        { field: "value", type: "quantitative", title: "$M" }
      ]
    },
    view: { stroke: null }
  };
}
