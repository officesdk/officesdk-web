import type {
  DocxMenuOptions,
  DocxToolbarOptions,
  DocxMenuEntryConfig,
  DocxToolbarFeatureButtonName,
  DocxMenuFeatureButtonName,
  DocxWindowScrollParams,
  DocxRangeValue,
  DocxWindow,
  DocxWindowScrollIntoViewOptions,
  DocxParagraphs
} from '@officesdk/editor-sdk-core/combine';
import type { EditorContent } from './editor';
import type {EditorOutline, EditorOutlineItem} from '@officesdk/editor-sdk-core/combine'

export enum DocumentExportType {
  Docx = 'docx',
  Pdf = 'pdf',
  Image = 'image',
}

export type DocumentMenuOptions = DocxMenuOptions
export type DocumentToolbarOptions = DocxToolbarOptions
export type DocumentMenuEntryConfig = DocxMenuEntryConfig
export type DocumentToolbarFeatureButtonName = DocxToolbarFeatureButtonName
export type DocumentMenuFeatureButtonName = DocxMenuFeatureButtonName
export type DocumentWindowScrollParams = DocxWindowScrollParams
export type DocumentRangeValue = DocxRangeValue
export type DocumentWindow = DocxWindow
export type DocumentWindowScrollIntoViewOptions = DocxWindowScrollIntoViewOptions
export type DocumentParagraphs = DocxParagraphs

/**
 * Document 远程调用的方法定义，
* 作为契约，用于统一约束客户端和服务端的接口。
 * 这里只有类型定义，不包含任何实现。
 */

export type DocumentMethods = {
  ready: () => Promise<void>;
  /**
   * 导出文档接口
   */
  export: (type: DocumentExportType) => Promise<void>;
  /**
   * 获取选区接口
   */
  getSelection: () => DocumentSelection;

  /**
   * 获取内容接口
   */
  getContent: () => EditorContent;

  /**
   * 获取当前文档的缩放接口
   */
  getZoom: () => DocumentZoom;

  /**
   * 文档目录集合接口，
   * 一个文档中可以存在多个目录，这个接口是用来管理文档中的所有目录的。
   */
  getTOCs: () => DocumentTOCs;

  /**
   *  传统文档目录大纲接口
   */
  getOutline: () => Omit<DocumentOutline, 'setVisible'>;

  /**
   * 文档窗口接口
   */
  getWindow: () => DocumentWindow;
  /**
   * 获取段落接口
   */
  getParagraphs: () => DocumentParagraphs;
  // TODO: 初始化流程控制，初始化各类异常
};

/**
 * 传统文档目录大纲项接口
 */
export type DocumentOutline = Omit<EditorOutline<{
  text: string;
}>, 'setVisible'>;

/**
 * 传统文档目录大纲项信息，用于描述传统文档中的目录项信息。
 */
export type DocumentOutlineItem = EditorOutlineItem<{
  text: string;
}>;

/**
 * 文档编辑器实例接口
 */
export interface DocumentEditor {
  readonly ready: () => Promise<void>;
  readonly export: (type: DocumentExportType) => Promise<void>;
  readonly selection: DocumentSelection;
  readonly zoom: DocumentZoom;
  readonly TOCs: DocumentTOCs;
  readonly outline: DocumentOutline;
  readonly window: DocumentWindow;
  readonly paragraphs: DocumentParagraphs;
}

export type DocumentSelection = {
  /**
   * 获取选区的区域范围，
   * 如果选区不存在，则返回 null。
   * @returns
   */
  getRange: (range?: DocumentRangeValue) => DocumentRange | null;

  /**
   * 设置选区的区域范围，
   * 设置后，选区会自动选中指定范围。
   * 如果设置为 null，则清空选区。
   * @param value 选区的区域范围
   */
  setRange: (range: DocumentRangeValue | null) => void;

  /**
   * 添加选区变化监听器，当选区发生变化时，会触发回调。
   * @param listener
   * @returns
   */
  addRangeListener: (listener: (range: DocumentRangeValue | null) => void) => void;

  /**
   * 获取文档全部选区范围。
   * @returns
   */
  getWholeRange: () => DocumentRange;
};

/**
 * 文档区域接口
 */
export interface DocumentRange {
  /**
   * 区域的开始位置，
   * 当文档发生变化后，区域的标识可能会失效。
   */
  readonly start: string;

  /**
   * 区域的结束位置，
   * 当文档发生变化后，区域的标识可能会失效。
   */
  readonly end: string;

  /**
   * 是否为光标。
   */
  readonly isCaret: boolean;

  /**
   * 获取该区域对应的纯文本信息。
   * @returns 返回纯文本字符串
   */
  getText: () => string;

  /**
   * 将区域内的文本设置为指定的文本，
   * 如果文本为空，则清空区域内的文本。
   * 如果区域为
   * @param text
   */
  setText: (text: string) => void;

  /**
   * 将区域中的内容以 HTML 格式返回。
   * @returns 返回 HTML 字符串
   */
  getHtml: () => string;
  /**
   * 将区域内的文本设置为指定的 HTML 内容，
   * 如果内容为空，则清空区域内的内容。
   * 如果区域为光标，则会将光标位置的文本替换为指定的 HTML 格式。
   * @param html 指定的 HTML 内容
   */
  setHtml: (html: string) => void;
  /**
   * 获取选区在当前屏幕上的位置信息
   * @returns
   */
  getBounding: () => DocumentRangeBounding | null;
}

/**
 * 记录区域在屏幕上的位置信息，
 * 包含了区域的四个边界值，以及开始和结束位置。
 * 因为选区是由多个矩形区域组成的，
 * 所以这里的位置信息是一个矩形区域的边界值。
 */
export interface DocumentRangeBounding {
  /**
   * 区域的左边界值，
   * 代表区域的最左边的 X 坐标，
   * 也就是区域的最左边的矩形区域的左边界值。
   */
  top: number;
  /**
   * 区域的上边界值，
   * 代表区域的最上边的 Y 坐标，
   * 也就是区域的最上边的矩形区域的上边界值。
   */
  right: number;
  /**
   * 区域的右边界值，
   * 代表区域的最右边的 X 坐标，
   * 也就是区域的最右边的矩形区域的右边界值。
   */
  bottom: number;
  /**
   * 区域的下边界值，
   * 代表区域的最下边的 Y 坐标，
   * 也就是区域的最下边的矩形区域的下边界值。
   */
  left: number;
  /**
   * 第一行的开始 X 坐标，
   */
  start: number;
  /**
   * 最后一行的结束 X 坐标，
   */
  end: number;
}

/**
 * 文档缩放接口
 */
export type DocumentZoom = {
  /**
   * 获取当前缩放比例。
   * @returns
   */
  getPercentage: () => number;
  /**
   * 设置缩放比例，
   * 有效范围 10 ~ 500。
   * @param percentage
   */
  setPercentage: (percentage: number) => void;
  /**
   * 设置自动缩放模式，
   * none: 不自动缩放，默认值。
   * window: 根据窗口宽度自动缩放，页面宽度随着窗口宽度变化而变化
   * page: 根据页面尺寸自动缩放，将页面缩放到可以完整显示的大小
   * @param mode 缩放模式，可以是 'window' 或 'page'
   */
  setFitMode: (mode: 'none' | 'window' | 'page') => void;
  /**
   * 获取当前缩放模式。
   * @returns
   */
  getFitMode: () => 'none' | 'window' | 'page';
  /**
   * 放大。
   */
  zoomIn: () => void;
  /**
   * 缩小。
   */
  zoomOut: () => void;
};

/**
 * 传统文档目录接口，
 * 表示文档中的所有目录，
 * 需要注意的是，DocumentTOCs 和 DocumentTocItem 是两个不同的对象，前者代表代表目录集合（一个文档可以插入多个目录），后者单个目录。
 */
export interface DocumentTOCs {
  /**
   * 获取所有目录列表
   * @returns 目录列表
   */
  getAll: () => Promise<DocumentTocItem[]>;
  /**
   * 获取某个目录
   * @param index 目录索引
   * @returns 目录
   */
  getOne: (index: number) => Promise<DocumentTocItem | null>;

  /**
   * 删除所有目录
   * @returns 删除结果
   */
  deleteAll: () => Promise<boolean>;

  /**
   * 删除某个目录
   * @param index
   * @returns 删除结果
   */
  deleteOne: (index: number) => Promise<boolean>;

  /**
   * 添加目录
   * @param options 目录选项
   * @returns 添加结果
   */
  add: (options: {
    /**
     * 添加目录的位置，
     * 默认为当前选区位置。
     * 如果添加失败，如：当前选区信息不正确，会返回 false。
     * 如果添加成功，会返回 true。
     */
    range?: string;
  }) => Promise<boolean>;
}

/**
 * 传统文档目录操作接口，
 * 可以调用接口更新目录和页码，
 * 也可以设置目录级别、页码等样式。
 */
export interface DocumentTocItem {
  /**
   * 获取目录信息
   * @returns 目录信息
   */
  getContent: () => Promise<DocumentTocContentItem[]>;
  /**
   * 添加目录信息改变时的监听器
   * @param listener 监听器
   * @returns 取消监听器的函数
   */
  addContentChangedListener: (listener: (content: DocumentTocContentItem[]) => void) => () => void;
  /**
   * 跳转到指定条目对应的正文位置
   * @param id 条目id
   * @returns 跳转是否成功
   */
  goto: (id: string) => Promise<boolean>;
  /**
   * 更新整个目录
   * @returns 是否更新成功
   */
  update: () => Promise<boolean>;
  /**
   * 仅更新页码
   * @returns 是否更新成功
   */
  updatePageNumbers: () => Promise<boolean>;
  /**
   * 设置目录层级
   * @param level 目录层级
   * @returns 是否设置成功
   */
  setLevel: (level: number) => Promise<boolean>;
}

/**
 * 通用目录项条目
 * @field id 目录项 ID
 * @field level 目录项层级
 * @field content 目录项内容
 */
export interface DocumentTocContentItem {
  /**
   * 目录项 ID
   */
  id: string;
  /**
   * 目录项层级
   */
  level: number;
  /**
   * 目录项内容
   */
  content: {
    /**
     * 目录项内容
     */
    text: string;
  };
}

export interface DocumentSDKOptions {
  /**
   * menu settings
   */
  menu?: {
    hidden?: boolean;
    disabled?: boolean;
    custom?: DocxMenuOptions['custom'];
    features?: DocxMenuOptions['features'];
  };
  /**
   * toolbar settings
   */
  toolbar?: {
    disabled?: boolean;
    hidden?: boolean;
    features?: DocxToolbarOptions['features'];
  };
}
