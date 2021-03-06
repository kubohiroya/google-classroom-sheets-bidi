/**
 * Create a Google Form by a sheet of Google Spreadsheet containing values representing Google Form contents.
 * @param sheet {Object} sheet of Google Spreadsheet containing values representing Google Form contents
 * @param [formTitle] {string} forms title (optional)
 * */

import {
  CheckboxGridItemObject,
  CheckboxItemObject,
  ChoiceObject,
  CorrectnessFeedbackObject,
  FormMetadataObject,
  FormSource,
  GeneralFeedbackObject,
  GridItemObject,
  ItemObject,
  ListItemObject,
  MultipleChoiceItemObject,
  QuizItemObject,
  QuizFeedbackObject,
  SurveyJsMatrixItemObject,
  SpreadsheetRow,
  SpreadsheetValues,
  SpreadsheetCell,
} from "./types";

type Sheet = GoogleAppsScript.Spreadsheet.Sheet;

const COL_INDEX = {
  TYPE: 0,
  META: {
    ID: 1,
    VERSION: 1,
    TITLE: 1,
    DESCRIPTION: 1,
    MESSAGE: 1,
    BOOLEAN: 1,
    URL: 1,
  },
  CHOICE: {
    VALUE: 1,
    IS_CORRECT_ANSWER: 2,
    PAGE_NAVIGATION_TYPE: 3,
    GOTO_PAGE_TITLE: 4,
  },
  GRID: {
    ROW_LABEL: 2,
    COL_LABEL: 2,
  },
  CHECKBOX_GRID: {
    ROW_LABEL: 2,
    COL_LABEL: 2,
  },
  ITEM: {
    TITLE: 1,
    HELP_TEXT: 2,
    Q: {
      REQUIRED: 3,
      SCALE: {
        LEFT_LABEL: 4,
        RIGHT_LABEL: 5,
        LOWER_BOUND: 6,
        UPPER_BOUND: 7,
      },
      DATE: {
        INCLUDES_YEAR: 4,
        POINTS: 5,
      },
      DATE_TIME: {
        INCLUDES_YEAR: 4,
        POINTS: 5,
      },
      TIME: {
        INCLUDES_YEAR: -1,
        POINTS: 5,
      },
      MULTIPLE_CHOICE: {
        SHOW_OTHER: 4,
        POINTS: 5,
      },
      CHECKBOX: {
        SHOW_OTHER: 4,
        POINTS: 5,
      },
      LIST: {
        SHOW_OTHER: -1, // DO NOT USE
        POINTS: 5,
      },
      TEXT: {
        POINTS: 5,
      },
      PARAGRAPH_TEXT: {
        POINTS: 5,
      },
    },
    VIDEO: {
      NA: 3,
      URL: 4,
      WIDTH: 5,
      ALIGNMENT: 6,
    },
    IMAGE: {
      NA: 3,
      URL: 4,
      WIDTH: 5,
      ALIGNMENT: 6,
    },
    PAGE_BREAK: {
      PAGE_NAVIGATION_TYPE: 3,
      GO_TO_PAGE_TITLE: 4,
    },
  },
  FEEDBACK: {
    TEXT: 1,
    URL: 1,
    DISPLAY_TEXT: 2,
  },
};

function createFormMetadata (
  values: SpreadsheetValues
): Partial<FormMetadataObject> {
  const metadataSrc: Record<string, SpreadsheetCell | string[]> = {};
  values.map((row: SpreadsheetRow) => {
    if (row.length >= 2) {
      const key = row[COL_INDEX.TYPE] as string;
      const value = row[1];
      if (key === "editors" && typeof value === "string") {
        metadataSrc[key] = value.split(",");
      } else {
        // if (typeof value === METADATA_TYPES[key])
        metadataSrc[key] = value;
      }
    }
  });
  return {} as Partial<FormMetadataObject>;
  //return (metadataSrc as unknown) as FormMetadataObject;
}

function createFormItemObjects(
  values: Array<Array<string | number | boolean | Date | null>>
): Array<ItemObject> {
  const itemObjects = Array<ItemObject>();
  values.forEach((row: Array<string | number | boolean | Date | null>) => {
    const command = row[COL_INDEX.TYPE] as string;
    if (command.charAt(0) === "#" || command === "comment") {
      return;
    }
    createFormItemObject(row, itemObjects);
  });
  return itemObjects;
}

function createFormItemObject (
  row: Array<string | number | boolean | Date | null>,
  itemObjects: Array<ItemObject>
): void {
  const type = row[COL_INDEX.TYPE] as string;
  if (type.startsWith("surveyJs:")) {
    if (type === "surveyJs:matrix") {
      const itemObject = {
        type: type,
        name: row[1] as string,
        columns: new Array<string>(),
        rows: new Array<{ value: string; text: string }>(),
        cells: {} as { [rowName: string]: { [columnName: string]: string } },
        points: {} as { [rowName: string]: { [columnName: string]: number } },
        _cellIndex: 0,
      };
      itemObjects.push({ ...itemObject, title: itemObject.name, helpText: "" });
    } else if (type === "surveyJs:matrix.column") {
      const item = (itemObjects[
        itemObjects.length - 1
      ] as unknown) as SurveyJsMatrixItemObject;
      item.columns.push(row[1] as string);
    } else if (type === "surveyJs:matrix.row") {
      const item = (itemObjects[
        itemObjects.length - 1
      ] as unknown) as SurveyJsMatrixItemObject;
      const text = row[1] as string;
      const value = `${item.name} [${
        (row[2] as string) || (row[1] as string)
      }]`;
      item.rows.push({ text: text, value: value });
      item.cells[value] = {};
      item.points[value] = {};
    } else if (type === "surveyJs:matrix.cell") {
      const item = (itemObjects[
        itemObjects.length - 1
      ] as unknown) as SurveyJsMatrixItemObject;
      const matrixRowIndex = Math.floor(item._cellIndex / item.columns.length);
      const matrixColumnIndex = item._cellIndex % item.columns.length;
      item.cells[item.rows[matrixRowIndex].value][
        item.columns[matrixColumnIndex]
      ] = row[1] as string;
      item.points[item.rows[matrixRowIndex].value][
        item.columns[matrixColumnIndex]
      ] = parseInt(row[2] as string);
      item._cellIndex++;
    }
  } else {
    const isQItem =
      type === "multipleChoice" ||
      type === "checkbox" ||
      type === "list" ||
      type === "date" ||
      type === "time" ||
      type === "dateTime" ||
      type === "duration" ||
      type === "grid" ||
      type === "checkboxGrid" ||
      type === "scale" ||
      type === "text" ||
      type === "paragraphText";
    const isOtherItem =
      type === "image" ||
      type === "video" ||
      type === "pageBreak" ||
      type === "sectionHeader";
    const isFeedback =
      type === "generalFeedback" ||
      type === "correctnessFeedback" ||
      type === "incorrectnessFeedback";
    const isFeedbackLink =
      type === "generalFeedback.link" ||
      type === "correctnessFeedback.link" ||
      type === "incorrectnessFeedback.link";
    const isChoice =
      type === "multipleChoice.choice" ||
      type === "checkbox.choice" ||
      type === "list.choice";

    if (isQItem || isOtherItem) {
      const itemObject = {
        type: type,
        title: row[COL_INDEX.ITEM.TITLE],
        helpText: row[COL_INDEX.ITEM.HELP_TEXT],
      } as ItemObject;

      if (isOtherItem) {
        if (type === "image") {
          const url = row[COL_INDEX.ITEM.IMAGE.URL];
          const width = row[COL_INDEX.ITEM.IMAGE.WIDTH];
          const alignment = row[COL_INDEX.ITEM.IMAGE.ALIGNMENT];
          const imageItemObject = {
            url,
            width,
            alignment,
            ...itemObject,
          };
          itemObjects.push(imageItemObject);
        } else if (type === "video") {
          const url = row[COL_INDEX.ITEM.VIDEO.URL];
          const width = row[COL_INDEX.ITEM.VIDEO.WIDTH];
          const alignment = row[COL_INDEX.ITEM.IMAGE.ALIGNMENT];
          const videoItemObject = {
            url,
            width,
            alignment,
            ...itemObject,
          };
          itemObjects.push(videoItemObject);
        } else if (type === "sectionHeader") {
          itemObjects.push({ ...itemObject });
        } else if (type === "pageBreak") {
          const gotoPageTitle = row[
            COL_INDEX.ITEM.PAGE_BREAK.GO_TO_PAGE_TITLE
          ] as string | null;
          const pageNavigationType = row[
            COL_INDEX.ITEM.PAGE_BREAK.PAGE_NAVIGATION_TYPE
          ] as string | null;
          if (gotoPageTitle) {
            const pageBreakObject = { gotoPageTitle, ...itemObject };
            itemObjects.push(pageBreakObject);
          } else if (pageNavigationType) {
            const pageBreakObject = { pageNavigationType, ...itemObject };
            itemObjects.push(pageBreakObject);
          } else {
            throw new Error();
          }
        }
      } else if (isQItem) {
        const qItemObject = itemObject as QuizItemObject;
        qItemObject.isRequired = row[COL_INDEX.ITEM.Q.REQUIRED] as boolean;
        if (type === "multipleChoice") {
          const showOther = row[COL_INDEX.ITEM.Q.MULTIPLE_CHOICE.SHOW_OTHER];
          const points = row[COL_INDEX.ITEM.Q.MULTIPLE_CHOICE.POINTS];
          const multipleChoiceItemObject = {
            showOther,
            points,
            choices: new Array<ChoiceObject>(),
            ...qItemObject,
          };
          itemObjects.push(multipleChoiceItemObject);
        } else if (type === "checkbox") {
          const showOther = row[COL_INDEX.ITEM.Q.CHECKBOX.SHOW_OTHER];
          const points = row[COL_INDEX.ITEM.Q.CHECKBOX.POINTS];
          const checkboxItemObject = {
            showOther,
            points,
            choices: new Array<ChoiceObject>(),
            ...qItemObject,
          };
          itemObjects.push(checkboxItemObject);
        } else if (type === "list") {
          const points = row[COL_INDEX.ITEM.Q.LIST.POINTS];
          const listItemObject = {
            points,
            choices: new Array<ChoiceObject>(),
            ...qItemObject,
          };
          itemObjects.push(listItemObject);
        } else if (type === "date") {
          const includesYear = row[COL_INDEX.ITEM.Q.DATE.INCLUDES_YEAR];
          const points = row[COL_INDEX.ITEM.Q.DATE.POINTS];
          const dateItemObject = { includesYear, points, ...qItemObject };
          itemObjects.push(dateItemObject);
        } else if (type === "time") {
          const points = row[COL_INDEX.ITEM.Q.TIME.POINTS];
          const dateItemObject = { points, ...qItemObject };
          itemObjects.push(dateItemObject);
        } else if (type === "dateTime") {
          const includesYear = row[COL_INDEX.ITEM.Q.DATE_TIME.INCLUDES_YEAR];
          const points = row[COL_INDEX.ITEM.Q.DATE_TIME.POINTS];
          const dateTimeItemObject = { includesYear, points, ...qItemObject };
          itemObjects.push(dateTimeItemObject);
        } else if (type === "grid") {
          const gridItemObject = {
            rows: new Array<string>(),
            columns: new Array<string>(),
            ...qItemObject,
          };
          itemObjects.push(gridItemObject);
        } else if (type === "checkboxGrid") {
          const checkboxGridItemObject = {
            rows: new Array<string>(),
            columns: new Array<string>(),
            ...qItemObject,
          };
          itemObjects.push(checkboxGridItemObject);
        } else if (type === "scale") {
          const leftLabel = row[COL_INDEX.ITEM.Q.SCALE.LEFT_LABEL];
          const lowerBound = row[COL_INDEX.ITEM.Q.SCALE.LOWER_BOUND];
          const rightLabel = row[COL_INDEX.ITEM.Q.SCALE.RIGHT_LABEL];
          const upperBound = row[COL_INDEX.ITEM.Q.SCALE.UPPER_BOUND];
          if (leftLabel && lowerBound && rightLabel && upperBound) {
            const scaleItemObject = {
              leftLabel,
              lowerBound,
              rightLabel,
              upperBound,
              ...qItemObject,
            };
            itemObjects.push(scaleItemObject);
          }
        } else if (type === "text") {
          const points = row[COL_INDEX.ITEM.Q.TEXT.POINTS];
          const textItemObject = { points, ...qItemObject };
          itemObjects.push(textItemObject);
        } else if (type === "paragraphText") {
          const points = row[COL_INDEX.ITEM.Q.PARAGRAPH_TEXT.POINTS];
          const paragraphTextItemObject = { points, ...qItemObject };
          itemObjects.push(paragraphTextItemObject);
        } else {
          throw new Error();
        }
      } else {
        throw new Error();
      }
    } else if (isFeedback) {
      const feedbackObject: QuizFeedbackObject = {
        text: row[COL_INDEX.FEEDBACK.TEXT] as string,
        links: [],
      };
      if (type === "generalFeedback") {
        const quizItem = (itemObjects[
          itemObjects.length - 1
        ] as unknown) as GeneralFeedbackObject;
        quizItem.generalFeedback = feedbackObject;
      } else if (type === "correctnessFeedback") {
        const quizItem = (itemObjects[
          itemObjects.length - 1
        ] as unknown) as CorrectnessFeedbackObject;
        quizItem.feedbackForCorrect = feedbackObject;
      } else if (type === "incorrectnessFeedback") {
        const quizItem = (itemObjects[
          itemObjects.length - 1
        ] as unknown) as CorrectnessFeedbackObject;
        quizItem.feedbackForIncorrect = feedbackObject;
      }
    } else if (isFeedbackLink) {
      const url = row[COL_INDEX.FEEDBACK.URL] as string;
      const displayText = row[COL_INDEX.FEEDBACK.DISPLAY_TEXT] as string;
      if (type === "generalFeedback.link") {
        const quizItem = (itemObjects[
          itemObjects.length - 1
        ] as unknown) as GeneralFeedbackObject;
        if (quizItem.generalFeedback) {
          quizItem.generalFeedback.links.push({ url, displayText });
        } else {
          throw new Error();
        }
      } else if (type === "correctnessFeedback.link") {
        const quizItem = (itemObjects[
          itemObjects.length - 1
        ] as unknown) as CorrectnessFeedbackObject;
        if (quizItem.feedbackForCorrect) {
          quizItem.feedbackForCorrect.links.push({ url, displayText });
        } else {
          throw new Error();
        }
      } else if (type === "incorrectnessFeedback.link") {
        const quizItem = (itemObjects[
          itemObjects.length - 1
        ] as unknown) as CorrectnessFeedbackObject;
        if (quizItem.feedbackForIncorrect) {
          quizItem.feedbackForIncorrect.links.push({ url, displayText });
        } else {
          throw new Error();
        }
      }
    } else if (isChoice) {
      const value = row[COL_INDEX.CHOICE.VALUE] as string;
      const choice =
        row[COL_INDEX.CHOICE.IS_CORRECT_ANSWER] === null
          ? ({
              value,
            } as ChoiceObject)
          : ({
              value,
              isCorrectAnswer: row[COL_INDEX.CHOICE.IS_CORRECT_ANSWER],
            } as ChoiceObject);
      if (type === "multipleChoice.choice") {
        const quizItem = (itemObjects[
          itemObjects.length - 1
        ] as unknown) as MultipleChoiceItemObject;
        const pageNavigationType = row[COL_INDEX.CHOICE.PAGE_NAVIGATION_TYPE];
        const gotoPageTitle = row[COL_INDEX.CHOICE.GOTO_PAGE_TITLE];
        if (typeof pageNavigationType === "string")
          choice.pageNavigationType = pageNavigationType as string;
        if (typeof gotoPageTitle === "string")
          choice.gotoPageTitle = gotoPageTitle as string;
        quizItem.choices.push(choice);
      } else if (type === "checkbox.choice") {
        const quizItem = (itemObjects[
          itemObjects.length - 1
        ] as unknown) as CheckboxItemObject;
        quizItem.choices.push(choice);
      } else if (type === "list.choice") {
        const quizItem = (itemObjects[
          itemObjects.length - 1
        ] as unknown) as ListItemObject;
        quizItem.choices.push(choice);
      } else {
        throw new Error();
      }
    } else if (type === "grid.row" || type === "grid.column") {
      const quizItem = (itemObjects[
        itemObjects.length - 1
      ] as unknown) as GridItemObject;
      if (type === "grid.row") {
        quizItem.rows.push(row[COL_INDEX.GRID.ROW_LABEL] as string);
      } else if (type === "grid.column") {
        quizItem.columns.push(row[COL_INDEX.GRID.COL_LABEL] as string);
      } else {
        throw new Error();
      }
    } else if (type === "checkboxGrid.row" || type === "checkboxGrid.column") {
      const quizItem = (itemObjects[
        itemObjects.length - 1
      ] as unknown) as CheckboxGridItemObject;
      if (type === "checkboxGrid.row") {
        quizItem.rows.push(row[COL_INDEX.CHECKBOX_GRID.ROW_LABEL] as string);
      } else if (type === "checkboxGrid.column") {
        quizItem.columns.push(row[COL_INDEX.CHECKBOX_GRID.COL_LABEL] as string);
      } else {
        throw new Error();
      }
    }
  }
};

export default function sheetToJson(
  sheet: Sheet,
  createdAt: number,
  updatedAt: number
): FormSource {
  const values = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();

  const metadata = {
    ...createFormMetadata(values),
    createdAt,
    updatedAt,
  } as FormMetadataObject;
  const items = createFormItemObjects(values);

  return {
    metadata,
    items,
  };
}
