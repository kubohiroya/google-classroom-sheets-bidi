import {
  CheckboxItemObject,
  DateItemObject,
  FormSource,
  GridItemObject,
  ItemObject,
  ListItemObject,
  MultipleChoiceItemObject,
  PageBreakItemObject,
  ParagraphTextItemObject,
  QuizItemObject,
  SectionHeaderItemObject,
  TextItemObject,
  FormItem,
} from "./types";

type SurveyJsItem = {
  type: string;
  name: string;
  html?: string;
};

type SurveyJsSourceObject = {
  title: string;
  completedHtml: string;
  pages: Array<SurveyJsPage>;
};
//Record<string,string>

type SurveyJsPage = {
  name: string;
  elements: Array<FormItem | SurveyJsItem>;
};

export function jsonToSurveyJs(formSrc: FormSource): SurveyJsSourceObject {
  const createPage = (
    pageNumber: number,
    options?: Record<string, string>
  ): SurveyJsPage => ({
    name: "page" + pageNumber,
    elements: [] as Array<FormItem | SurveyJsItem>,
    ...options,
  });

  const formSrcItemToSurveyJsItem = (item: ItemObject) => {
    if (item.type === "surveyJs:matrix") {
      return {
        ...item,
        type: "matrix",
        name: item.title,
      };
    } else if (item.type === "checkbox") {
      const itemObject = item as ItemObject &
        QuizItemObject &
        CheckboxItemObject;
      return {
        type: "checkbox",
        name: itemObject.title,
        title: itemObject.title,
        tooltip: itemObject.helpText,
        isRequired: itemObject.isRequired,
        choices: itemObject.choices,
      };
    } else if (item.type === "checkboxGrid") {
      throw new Error("Not implemented: " + item.type);
    } else if (item.type === "date") {
      const itemObject = item as ItemObject & DateItemObject;
      return {
        type: "datepicker",
        name: itemObject.title,
        title: itemObject.title,
        tooltip: itemObject.helpText,
        inputType: "date",
        dateFormat: "yyyy-mm-dd",
        isRequired: true,
      };
    } else if (item.type === "dateTime") {
      const itemObject = item as ItemObject & DateItemObject;
      return {
        type: "datepicker",
        name: itemObject.title,
        title: itemObject.title,
        tooltip: itemObject.helpText,
        inputType: "datetime",
        dateFormat: "yyyy-mm-dd HH:MM",
        isRequired: true,
      };
    } else if (item.type === "duration") {
      throw new Error("Not implemented: " + item.type);
    } else if (item.type === "grid") {
      const itemObject = item as ItemObject & GridItemObject;
      return {
        type: "matrix",
        name: itemObject.title,
        title: itemObject.title,
        tooltip: itemObject.helpText,
        isRequired: itemObject.isRequired,
        columns: itemObject.columns.map((c) => ({ value: c, text: c })),
        row: itemObject.rows.map((c) => ({ value: c, text: c })),
      };
    } else if (item.type === "image") {
      throw new Error("Not implemented: " + item.type);
    } else if (item.type === "list") {
      const itemObject = item as ItemObject & ListItemObject;
      return {
        type: "dropdown",
        name: itemObject.title,
        title: itemObject.title,
        tooltip: itemObject.helpText,
        isRequired: itemObject.isRequired,
        choices: itemObject.choices,
      };
    } else if (item.type === "multipleChoice") {
      const itemObject = item as ItemObject & MultipleChoiceItemObject;
      return {
        type: "radiogroup",
        name: itemObject.title,
        title: itemObject.title,
        tooltip: itemObject.helpText,
        isRequired: itemObject.isRequired,
        choices: itemObject.choices,
      };
    } else if (item.type === "paragraphText") {
      const itemObject = item as ItemObject & ParagraphTextItemObject;
      return {
        type: "comment",
        name: itemObject.title,
        title: itemObject.title,
        tooltip: itemObject.helpText,
        isRequired: itemObject.isRequired,
      };
    } else if (item.type === "scale") {
      throw new Error("Not implemented: " + item.type);
    } else if (item.type === "sectionHeader") {
      const itemObject = item as SectionHeaderItemObject;
      return {
        type: "html",
        name: itemObject.title,
        tooltip: itemObject.helpText,
        html: `<h3>${itemObject.title}</h3><h4>${itemObject.helpText}</h4>`,
      };
    } else if (item.type === "text") {
      const itemObject = item as ItemObject & TextItemObject;
      return {
        type: "text",
        name: itemObject.title,
        title: itemObject.title,
        tooltip: itemObject.helpText,
        isRequired: itemObject.isRequired,
      };
    } else if (item.type === "time") {
      throw new Error("Not implemented: " + item.type);
    } else {
      throw new Error("Invalid type: " + item.type);
    }
  };

  const pages = [createPage(1)] as Array<SurveyJsPage>;

  formSrc.items.forEach((item) => {
    const itemObject = item as ItemObject;
    if (itemObject.type === "pageBreak") {
      const itemObject = item as PageBreakItemObject;
      const page = createPage(pages.length + 1, {
        title: itemObject.title,
      });
      page.elements.push({
        type: "html",
        name: itemObject.title,
        html: `<p>${itemObject.helpText}</p>`,
      });
      pages.push(page);
    } else {
      try {
        const surveyJsItem = formSrcItemToSurveyJsItem(itemObject);
        pages[pages.length - 1].elements.push(surveyJsItem);
      } catch (ignore) {
        Logger.log("ignore: " + itemObject.type);
      }
    }
  });

  return {
    title: formSrc.metadata.title,
    completedHtml: `<h3>${
      formSrc.metadata.confirmationMessage || "入力内容を送信しました"
    }</h3>`,
    pages,
  } as SurveyJsSourceObject;
}
