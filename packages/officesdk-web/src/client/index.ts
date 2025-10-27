export { createSDK } from './create';
export type { SDKSettings } from './create';
export type {
  CreateOptions,
  DocumentCreateOptions,
  PresentationCreateOptions,
  SpreadsheetCreateOptions,
  DBTableCreateOptions,
  LiteDocCreateOptions,
  PdfCreateOptions,
  BaseCreateOptions,
  OfficeSDK,
  OfficeSDKMap,
  OfficeDocumentSDK,
  OfficeSpreadSheetSDK,
  OfficePresentationSDK,
  OfficePdfSDK,
  OfficeDBTableSDK,
  OfficeLiteDocSDK,
} from './create';

export type { createContentFacade } from './editor/content';

// 公共部分
export { FileType } from '../shared/file';

export type { LTDocFacade } from './ltdoc';
export type { DatabaseTableFacade, DatabaseTableSettings } from './dbtable';

// 文档
export type { DocumentFacade, DocumentSettings, DocxMenuOptions, DocxToolbarOptions, DocumentParagraphs } from './document';

// 表格
export type { SpreadsheetFacade, SpreadsheetSettings } from './spreadsheet';

// 幻灯片
export type { PresentationFacade, PresentationSettings } from './presentation';

// PDF
export type { PdfFacade } from './pdf';
