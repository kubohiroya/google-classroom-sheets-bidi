type Item = GoogleAppsScript.Forms.Item;
type QuizFeedback = GoogleAppsScript.Forms.QuizFeedback;
type MultipleChoiceItem = GoogleAppsScript.Forms.MultipleChoiceItem;
type CheckboxItem = GoogleAppsScript.Forms.CheckboxItem;
type CheckboxGridItem = GoogleAppsScript.Forms.CheckboxGridItem;
type DateItem = GoogleAppsScript.Forms.DateItem;
type DateTimeItem = GoogleAppsScript.Forms.DateTimeItem;
type DurationItem = GoogleAppsScript.Forms.DurationItem;
type GridItem = GoogleAppsScript.Forms.GridItem;
type ImageItem = GoogleAppsScript.Forms.ImageItem;
type ListItem = GoogleAppsScript.Forms.ListItem;
type PageBreakItem = GoogleAppsScript.Forms.PageBreakItem;
type ScaleItem = GoogleAppsScript.Forms.ScaleItem;
type SectionHeaderItem = GoogleAppsScript.Forms.SectionHeaderItem;
type TextItem = GoogleAppsScript.Forms.TextItem;
type Alignment = GoogleAppsScript.Forms.Alignment;
type Choice = GoogleAppsScript.Forms.Choice;

export type FormItem =
  | CheckboxItem
  | CheckboxGridItem
  | DateItem
  | DateTimeItem
  | DurationItem
  | GridItem
  | ImageItem
  | ListItem
  | MultipleChoiceItem
  | PageBreakItem
  | ScaleItem
  | SectionHeaderItem
  | TextItem;

export interface FormMetadataObject {
  createdAt: number;
  updatedAt: number;
  quiz?: boolean;
  allowResponseEdits?: boolean;
  collectEmail?: boolean;
  description: string;
  acceptingResponses?: boolean;
  publishingSummary?: boolean;
  confirmationMessage?: string;
  customClosedFormMessage?: string;
  limitOneResponsePerUser?: boolean;
  progressBar?: boolean;
  editUrl: string;
  editors: string[];
  id: string;
  publishedUrl: string;
  shuffleQuestions?: boolean;
  summaryUrl?: string;
  title: string;
  requiresLogin?: boolean;
  linkToRespondAgain?: boolean;
  destinationId?: string;
  destinationType?: string;
}

export interface FormSource {
  metadata: FormMetadataObject;
  items: Array<
    ItemObject &
      (
        | QuizItemObjectTypeSet
        | GeneralFeedbackObject
        | CorrectnessFeedbackObject
        | OtherItemObject
      )
  >;
}

export const GOOGLE_FORMS_TYPES = {
  checkbox: FormApp.ItemType.CHECKBOX,
  checkboxGrid: FormApp.ItemType.CHECKBOX_GRID,
  date: FormApp.ItemType.DATE,
  dateTime: FormApp.ItemType.DATETIME,
  duration: FormApp.ItemType.DURATION,
  grid: FormApp.ItemType.GRID,
  image: FormApp.ItemType.IMAGE,
  list: FormApp.ItemType.LIST,
  multipleChoice: FormApp.ItemType.MULTIPLE_CHOICE,
  pageBreak: FormApp.ItemType.PAGE_BREAK,
  paragraphText: FormApp.ItemType.PARAGRAPH_TEXT,
  scale: FormApp.ItemType.SCALE,
  sectionHeader: FormApp.ItemType.SECTION_HEADER,
  text: FormApp.ItemType.TEXT,
  time: FormApp.ItemType.TIME,
};

export const SERVEYJS_ITEM_TYPES = {
  MATRIX: "matrix",
};

export const TYPE_NAMES = {
  [FormApp.ItemType.CHECKBOX]: "checkbox",
  [FormApp.ItemType.CHECKBOX_GRID]: "checkboxGrid",
  [FormApp.ItemType.DATE]: "date",
  [FormApp.ItemType.DATETIME]: "dateTime",
  [FormApp.ItemType.DURATION]: "duration",
  [FormApp.ItemType.GRID]: "grid",
  [FormApp.ItemType.IMAGE]: "image",
  [FormApp.ItemType.LIST]: "list",
  [FormApp.ItemType.MULTIPLE_CHOICE]: "multipleChoice",
  [FormApp.ItemType.PAGE_BREAK]: "pageBreak",
  [FormApp.ItemType.PARAGRAPH_TEXT]: "paragraphText",
  [FormApp.ItemType.SCALE]: "scale",
  [FormApp.ItemType.SECTION_HEADER]: "sectionHeader",
  [FormApp.ItemType.TEXT]: "text",
  [FormApp.ItemType.TIME]: "time",
  [SERVEYJS_ITEM_TYPES.MATRIX]: "surveyJs:matrix",
};

export interface SurveyJsMatrixItemObject {
  type: "surveyJs:matrix";
  name: string;
  columns: Array<string>;
  rows: Array<{ value: string; text: string }>;
  cells: { [rowName: string]: { [columnName: string]: string } };
  points: { [rowName: string]: { [columnName: string]: number } };
  _cellIndex: number;
}

export interface BlobObject {
  name: string;
  contentType: string;
  dataAsString: string;
}

export interface QuizFeedbackObject {
  text: string;
  links: Array<{ url: string; displayText?: string }>;
}

export interface GeneralFeedbackObject {
  points: number;
  generalFeedback?: QuizFeedbackObject;
}

export interface CorrectnessFeedbackObject {
  points: number;
  feedbackForCorrect?: QuizFeedbackObject;
  feedbackForIncorrect?: QuizFeedbackObject;
}

export interface ItemObject {
  type: string;
  // id: number;
  // index: number;
  title: string;
  helpText: string;
}

export interface QuizItemObject extends ItemObject {
  isRequired: boolean;
}

export type CheckboxGridItemObject = GridItemObject;

export interface HasChoicesObject {
  choices: ChoiceObject[];
}

export interface HasOtherOptionObject {
  hasOtherOption: boolean;
}

export interface PageNavigationObject {
  gotoPageTitle?: string;
  pageNavigationType?: string;
}

export interface ChoiceObject extends PageNavigationObject {
  value: string;
  isCorrectAnswer?: boolean;
}

export interface PageBreakItemObject extends ItemObject, PageNavigationObject {}

export interface CheckboxItemObject
  extends QuizItemObject,
    CorrectnessFeedbackObject,
    HasChoicesObject,
    HasOtherOptionObject {}

export interface IncludesYearObject extends QuizItemObject {
  includesYear: boolean;
}
export interface DateItemObject
  extends IncludesYearObject,
    GeneralFeedbackObject {}
export interface DateTimeItemObject
  extends DateItemObject,
    GeneralFeedbackObject {}
export interface DurationItemObject
  extends QuizItemObject,
    GeneralFeedbackObject {}

export interface GridItemObject extends QuizItemObject {
  rows: string[];
  columns: string[];
}

export interface ImageItemObject extends ItemObject {
  image?: BlobObject;
  url: string;
  alignment: string;
  width: number;
}

export interface ListItemObject
  extends QuizItemObject,
    HasChoicesObject,
    CorrectnessFeedbackObject {}

export interface MultipleChoiceItemObject
  extends QuizItemObject,
    CorrectnessFeedbackObject,
    HasChoicesObject,
    HasOtherOptionObject {}

export interface ParagraphTextItemObject
  extends QuizItemObject,
    GeneralFeedbackObject {}

export interface ScaleItemObject extends QuizItemObject, GeneralFeedbackObject {
  leftLabel: string;
  lowerBound: number;
  rightLabel: string;
  upperBound: number;
}

export type SectionHeaderItemObject = ItemObject;

export interface TextItemObject extends QuizItemObject, GeneralFeedbackObject {}

export interface TimeItemObject extends QuizItemObject, GeneralFeedbackObject {}

export interface VideoItemObject extends ItemObject {
  videoUrl: string;
  alignment: string;
  width: number;
}

export type OtherItemObject =
  | ImageItemObject
  | VideoItemObject
  | PageBreakItemObject
  | SectionHeaderItemObject;

export type QuizItemObjectTypeSet =
  | TextItemObject
  | ParagraphTextItemObject
  | CheckboxItemObject
  | MultipleChoiceItemObject
  | DateItemObject
  | DateTimeItemObject
  | DurationItemObject
  | TimeItemObject
  | GridItemObject
  | CheckboxGridItemObject
  | ListItemObject
  | ScaleItemObject;

export interface QuizItem extends Item {
  isRequired(): boolean;
  setRequired(value: boolean): void;
}

export interface ShowOtherOptionItem extends QuizItem {
  showOtherOption(enabled: boolean): void;
}

export interface GeneralFeedbackItem extends QuizItem {
  setPoints(points: number): void;
  getPoints(): number;
  setGeneralFeedback(feedback: QuizFeedback): void;
  getGeneralFeedback(): QuizFeedback;
}

export interface CorrectnessFeedbackItem extends QuizItem {
  setPoints(points: number): void;
  getPoints(): number;
  setFeedbackForCorrect(feedback: QuizFeedback): void;
  setFeedbackForIncorrect(feedback: QuizFeedback): void;
  getFeedbackForCorrect(): QuizFeedback;
  getFeedbackForIncorrect(): QuizFeedback;
}

export interface HasOtherOptionsItem extends CorrectnessFeedbackItem {
  hasOtherOption(): boolean;
  getChoices(): Choice[];
}

export interface IncludesYearItem extends QuizItem {
  includesYear(): boolean;
  setIncludesYear(enableYear: boolean): void;
}

export interface AddMultipleChoiceItem extends QuizItem {
  addMultipleChoiceItem(): MultipleChoiceItem;
}

export function getAlignment(value: string): Alignment {
  return value === "LEFT"
    ? FormApp.Alignment.LEFT
    : value === "CENTER"
    ? FormApp.Alignment.CENTER
    : value === "RIGHT"
    ? FormApp.Alignment.RIGHT
    : FormApp.Alignment.RIGHT;
}

export type SpreadsheetCell = string | number | boolean | Date | undefined;
export type SpreadsheetRow = Array<SpreadsheetCell>;
export type SpreadsheetValues = Array<SpreadsheetRow>;
