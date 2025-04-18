import type { EditorOutline, EditorOutlineItem } from './editor';

/**
 * Pdf 远程调用的方法定义，
 * 作为契约，用于统一约束客户端和服务端的接口。
 * 这里只有类型定义，不包含任何实现。
 */
export type PdfMethods = {
  /**
   * 获取页面操作接口
   */
  getPages: () => PdfPages;

  /**
   * 获取选区接口
   */
  getSelection: () => PdfSelection;

  /**
   * 获取目录接口
   */
  getOutline: () => PdfOutline;
};

/**
 * PDF 编辑器目录项接口，用于访问 pdf 文件中目录信息。
 */
export type PdfOutline = EditorOutline<{
  text: string;
}>;

/**
 * PDF 编辑器目录项信息，用于描述 PDF 编辑器中的目录项信息。
 */
export type PdfOutlineItem = EditorOutlineItem<{
  text: string;
}>;

/**
 * PDF 编辑器实例接口
 */
export interface PdfEditor {
  readonly pages: PdfPages;
  readonly selection: PdfSelection;
  readonly outline: PdfOutline;
}

/**
 * PDF 页面集合操作接口
 */
export interface PdfPages {
  /**
   * 获取当前页码
   * @returns
   */
  getCurrentPageNumber: () => number;

  /**
   * 设置当前页码并跳转至对应页面
   * @param page
   */
  setCurrentPage: (page: number) => void;

  /**
   * 获取总页数
   * @returns
   */
  getPagesCount: () => number;

  /**
   * 获取指定页面
   */
  getPage: (page: number) => Promise<PdfPage>;
}

/**
 * PDF 页面实例
 */
export interface PdfPage {
  /**
   * 获取当前页码
   */
  readonly pageNumber: number;

  /**
   * 获取当前页面的大小
   * @returns
   */
  getPageSize: () => {
    width: number;
    height: number;
  };
}

/**
 * PDF 选区接口
 */
export interface PdfSelection {
  /**
   * 获取选区的区域范围，
   * 如果没有指定范围，则返回当前选区的范围。
   * 如果指定了范围，则返回指定范围的选区。
   * 如果选区不存在，则返回 null。
   * @returns
   */
  getRange: (value?: PdfRangeValue) => PdfRange | null;

  /**
   * 设置选区的区域范围，
   * 设置后，选区会自动选中指定范围。
   * 如果设置为 null，则清空选区。
   * @param value
   */
  setRange: (value: PdfRangeValue | null) => void;

  /**
   * 添加选区变化监听器，当选区发生变化时，会触发回调。
   * @param listener
   * @returns
   */
  addRangeListener: (listener: (value: PdfRangeValue) => void) => void;
}

/**
 * PDF 区域对象
 */
export interface PdfRange {
  /**
   * 区域的开始位置
   */
  readonly start: string;

  /**
   * 区域的结束位置
   */
  readonly end: string;

  /**
   * 获取该区域对应的纯文本信息。
   * @returns
   */
  getText: () => string;

  /**
   * 将区域中的内容以 HTML 格式返回。
   * @returns
   */
  getHtml: () => string;
}

/**
 * 可用于描述一个 PDF 区域的信息。
 */
export interface PdfRangeValue {
  start: string;
  end: string;
}
