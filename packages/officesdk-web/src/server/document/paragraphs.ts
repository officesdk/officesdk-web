import type { DocumentParagraphs } from "../../shared";
import { createDocumentRangeProxy } from "./range";

export function createDocumentParagraphsProxy(paragraphs: DocumentParagraphs): DocumentParagraphs {
  return {
    getAll: async () => {
      const paragraphList = await paragraphs.getAll();
      return paragraphList.map(para => ({
        getRange: async() => {
          const range = await para.getRange();
          return createDocumentRangeProxy(range)
        }
      }))
    },
    getOne: async (index: number) => {
      const para = await paragraphs.getOne(index);
      if (!para) {
        return null;
      }
      return {
        getRange: async() => {
          const range = await para.getRange();
          return createDocumentRangeProxy(range)
        }
      }
    },
    count: async () => {
      return await paragraphs.count();
    }
  }
}