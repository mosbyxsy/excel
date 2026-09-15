/**
 * 站点部署配置。
 *
 * 这是唯一需要部署人员按实际环境修改的文件。页面会从
 * window.EXCEL_VIEWER_CONFIG 读取预置文件列表和虚拟滚动参数。
 * 修改后直接刷新浏览器即可生效，不需要运行任何构建命令。
 */
window.EXCEL_VIEWER_CONFIG = {
  /**
   * 预置的远程文件列表。
   *
   * 每一项包含：
   * - id：文件的稳定唯一标识；
   * - name：显示在下拉列表中的名称；
   * - url：文件地址，可以是同源相对地址或允许 CORS 的完整 URL；
   * - type：可选，支持 xlsx、xlsm、xls、csv。省略时根据地址和文件签名判断。
   * - action：可选；设置为 "open" 时页面启动后默认打开该文件。
   * - autoFit：可选；true 表示打开后默认按原列宽比例适应可视区域，默认 false。
   *
   * 通过 ?config= 加载的远程 JSON 也使用 { files: [] } 结构和相同字段。
   * 若合并后的列表中有多个 action: "open"，页面只打开排列在最前面的一个。
   *
   * 示例：
   * { id: "report", name: "月度报表", url: "/files/report.xlsx", type: "xlsx", autoFit: false }
   */
  files: [{
    id: "local_id_1",
    name: "Webstrom和VScode快捷键",
    url: "/excel/Webstrom和VScode快捷键.xlsx",
    type: "xlsx",
    action: "open",
    autoFit: false
  }, {
    id: "local_id_2",
    name: "Windows和Mac平台Edge全量快捷键",
    url: "/excel/Windows和Mac平台Edge全量快捷键.xlsx",
    type: "xlsx",
    autoFit: false
  }, {
    id: "local_id_3",
    name: "Windows快捷键总表",
    url: "/excel/Windows快捷键总表.xlsx",
    type: "xlsx",
    autoFit: false
  }, {
    id: "local_id_4",
    name: "Windows平台Edge全量快捷键",
    url: "/excel/Windows平台Edge全量快捷键.xlsx",
    type: "xlsx",
    autoFit: false
  }, {
    id: "local_id_5",
    name: "Mac快捷键总表",
    url: "/excel/Mac快捷键总表.xlsx",
    type: "xlsx",
    autoFit: false
  }, {
    id: "local_id_6",
    name: "Mac平台Edge全量快捷键",
    url: "/excel/Mac平台Edge全量快捷键.xlsx",
    type: "xlsx",
    autoFit: false
  }, , {
    id: "local_id_7",
    name: "Windows和Mac快捷键",
    url: "/excel/Windows和Mac快捷键.xlsx",
    type: "xlsx",
    autoFit: true
  }, {
    id: "local_id_8",
    name: "超级表测试",
    url: "/excel/超级表测试.xlsx",
    type: "xlsx",
    action: "open",
    autoFit: false
  }],

  /** 页面打开后自动加载的预置文件 id；null 表示不自动加载。 */
  defaultFileId: null,

  /** 有效行数超过该值后启用行虚拟化，避免一次创建过多 DOM 节点。 */
  virtualizationThreshold: 500,

  /** 虚拟列表在可视区域上下额外保留的行数，用于减少快速滚动时的空白闪烁。 */
  overscanRows: 10
};
