type Form = GoogleAppsScript.Forms.Form;
type Blob = GoogleAppsScript.Base.Blob;
type QuizFeedback = GoogleAppsScript.Forms.QuizFeedback;
type Alignment = GoogleAppsScript.Forms.Alignment;
type Choice = GoogleAppsScript.Forms.Choice;
type PageBreakItem = GoogleAppsScript.Forms.PageBreakItem;
type Item = GoogleAppsScript.Forms.Item;
type CheckboxItem = GoogleAppsScript.Forms.CheckboxItem;
type CheckboxGridItem = GoogleAppsScript.Forms.CheckboxGridItem;
type DateItem = GoogleAppsScript.Forms.DateItem;
type DateTimeItem = GoogleAppsScript.Forms.DateTimeItem;
type DurationItem = GoogleAppsScript.Forms.DurationItem;
type GridItem = GoogleAppsScript.Forms.GridItem;
type MultipleChoiceItem = GoogleAppsScript.Forms.MultipleChoiceItem;
type ParagraphTextItem = GoogleAppsScript.Forms.ParagraphTextItem;
type ScaleItem = GoogleAppsScript.Forms.ScaleItem;
type TextItem = GoogleAppsScript.Forms.TextItem;
type ImageItem = GoogleAppsScript.Forms.ImageItem;
type VideoItem = GoogleAppsScript.Forms.VideoItem;
const DestinationType = FormApp.DestinationType;
import {
  BlobObject,
  ChoiceObject,
  CorrectnessFeedbackItem,
  CorrectnessFeedbackObject,
  FormItem,
  FormMetadataObject,
  FormSource,
  GeneralFeedbackItem,
  GeneralFeedbackObject,
  OtherItemObject,
  QuizItemObjectTypeSet,
  QuizFeedbackObject,
  TYPE_NAMES,
  ItemObject,
  DateItemObject,
  DateTimeItemObject,
  TimeItemObject,
  DurationItemObject,
  CheckboxGridItemObject,
  GridItemObject,
  ParagraphTextItemObject,
  ScaleItemObject,
  TextItemObject,
  ImageItemObject,
  VideoItemObject,
  PageBreakItemObject,
  MultipleChoiceItemObject,
  CheckboxItemObject,
  ListItemObject,
  SectionHeaderItemObject,
} from "./types";
type ListItem = GoogleAppsScript.Forms.ListItem;
type TimeItem = GoogleAppsScript.Forms.TimeItem;

/**
 * Convert a Google Form object into JSON Object.
 * @returns JSON object of forms data
 * @param form {Form} a Google Form Object
 * */

const getFormMetadata = (
  form: Form,
  createdAt: number,
  updatedAt: number
): FormMetadataObject => {
  const metadata: FormMetadataObject = {
    quiz: form.isQuiz(),
    collectEmail: form.collectsEmail(),
    progressBar: form.hasProgressBar(),
    linkToRespondAgain: form.hasRespondAgainLink(),
    limitOneResponsePerUser: form.hasLimitOneResponsePerUser(),
    requiresLogin: form.requiresLogin(),
    allowResponseEdits: form.canEditResponse(),
    acceptingResponses: form.isAcceptingResponses(),
    publishingSummary: form.isPublishingSummary(),
    confirmationMessage: form.getConfirmationMessage(),
    customClosedFormMessage: form.getCustomClosedFormMessage(),
    description: form.getDescription(),
    editUrl: form.getEditUrl(),
    editors: form.getEditors().map((user) => user.getEmail()),
    id: form.getId(),
    publishedUrl: form.getPublishedUrl(),
    shuffleQuestions: form.getShuffleQuestions(),
    summaryUrl: form.getSummaryUrl(),
    title: form.getTitle(),
    createdAt,
    updatedAt,
  };

  try {
    const destinationId = form.getDestinationId();
    const destinationType = getDestinationTypeString(form.getDestinationType());
    metadata.destinationId = destinationId;
    metadata.destinationType = destinationType;
  } catch (ignore) {
    // do nothing
  }

  return metadata;
};

export default function formToJson(
  form: Form,
  createdAt: number,
  updatedAt: number
): FormSource {
  return {
    metadata: getFormMetadata(form, createdAt, updatedAt),
    items: form.getItems().map(itemToObject),
  };
}

function blobToJson(blob: Blob): BlobObject {
  return {
    name: blob.getName(),
    contentType: blob.getContentType(),
    dataAsString: blob.getDataAsString(),
  };
}

function choiceToJson(choice: Choice): ChoiceObject {
  const navigation = choice.getGotoPage()
    ? {
        gotoPageTitle: choice.getGotoPage().getTitle(),
        pageNavigationType: getPageNavigationTypeString(
          choice.getPageNavigationType()
        ),
      }
    : {};

  return {
    value: choice.getValue(),
    isCorrectAnswer: choice.isCorrectAnswer(),
    ...navigation,
  };
}

function feedbackToJson(
  type: string,
  feedback: QuizFeedback
): QuizFeedbackObject | null {
  if (feedback) {
    return {
      text: feedback.getText(),
      links: feedback.getLinkUrls().map((linkUrl) => ({ url: linkUrl })),
    };
  } else {
    return null;
  }
}

function getDestinationTypeString(type: number) {
  switch (type) {
    case DestinationType.SPREADSHEET:
      return "SPREADSHEET";
    default:
      throw new Error("INVALID_DESTINATION");
  }
}

function getAlignmentString(alignment: Alignment) {
  switch (alignment) {
    case FormApp.Alignment.LEFT:
      return "LEFT"; // FormApp.Alignment.LEFT.toString(); //;
    case FormApp.Alignment.RIGHT:
      return "RIGHT";
    case FormApp.Alignment.CENTER:
      return "CENTER";
    default:
      throw new Error("INVALID_ALIGNMENT");
  }
}

function getPageNavigationTypeString(type: number) {
  switch (type) {
    case FormApp.PageNavigationType.CONTINUE:
      return "CONTINUE";
    case FormApp.PageNavigationType.GO_TO_PAGE:
      return "GO_TO_PAGE";
    case FormApp.PageNavigationType.RESTART:
      return "RESTART";
    case FormApp.PageNavigationType.SUBMIT:
      return "SUBMIT";
    default:
      throw new Error("INVALID_NAVIGATION");
  }
}

/*
function getTypedItem(item: Item): FormItem{
  switch (item.getType()) {
    case FormApp.ItemType.CHECKBOX:
      return item.asCheckboxItem();
    case FormApp.ItemType.CHECKBOX_GRID:
      return item.asCheckboxGridItem();
    case FormApp.ItemType.DATE:
      return item.asDateItem();
    case FormApp.ItemType.DATETIME:
      return item.asDateTimeItem();
    case FormApp.ItemType.TIME:
      return item.asTimeItem();
    case FormApp.ItemType.DURATION:
      return item.asDurationItem();
    case FormApp.ItemType.GRID:
      return item.asGridItem();
    case FormApp.ItemType.IMAGE:
      return item.asImageItem();
    case FormApp.ItemType.LIST:
      return item.asListItem();
    case FormApp.ItemType.MULTIPLE_CHOICE:
      return item.asMultipleChoiceItem();
    case FormApp.ItemType.PAGE_BREAK:
      return item.asPageBreakItem();
    case FormApp.ItemType.PARAGRAPH_TEXT:
      return item.asParagraphTextItem();
    case FormApp.ItemType.SCALE:
      return item.asScaleItem();
    case FormApp.ItemType.SECTION_HEADER:
      return item.asSectionHeaderItem();
    case FormApp.ItemType.TEXT:
      return item.asTextItem();
    default:
      throw new Error("TypeError=" + item.getType());
  }
}*/

function createItemObject(item: Item): ItemObject {
  return {
    type: TYPE_NAMES[item.getType()],
    // id: item.getId(),
    // index: item.getIndex(),
    title: item.getTitle(),
    helpText: item.getHelpText(),
  };
}

// isRequired: item.isRequired(),

function createCorrectnessFeedbackObject(
  item: FormItem
): CorrectnessFeedbackObject {
  const typedItem = item as CorrectnessFeedbackItem;
  const feedbackForCorrect = typedItem.getFeedbackForCorrect ? feedbackToJson(
    "correctnessFeedback",
    typedItem.getFeedbackForCorrect()
  ) : undefined;
  const feedbackForIncorrect = typedItem.getFeedbackForIncorrect ? feedbackToJson(
    "correctnessFeedback",
    typedItem.getFeedbackForIncorrect()
  ): undefined;
  const points = typedItem.getPoints? typedItem.getPoints() : undefined;
  return {
    points,
    feedbackForCorrect,
    feedbackForIncorrect,
  } as CorrectnessFeedbackObject;
}

function createGeneralFeedbackObject(quizFeedback: QuizFeedback, points: number): GeneralFeedbackObject {
  const generalFeedback = feedbackToJson(
    "generalFeedback",
    quizFeedback,
  );
  return {
    points,
    ...(generalFeedback ? { ...generalFeedback } : {}),
  };
}

function itemToObject(
  item: Item
):
  | (QuizItemObjectTypeSet &
      (GeneralFeedbackObject | CorrectnessFeedbackObject))
  | OtherItemObject {
  const typedItemObject = createItemObject(item);

  Logger.log(item.getIndex()+"  "+item.getType().toString()+" "+item.getTitle());

  switch (item.getType()) {
    case FormApp.ItemType.CHECKBOX:
    case FormApp.ItemType.MULTIPLE_CHOICE:
      return (() => {
        const typedItem = (item.getType() === FormApp.ItemType.CHECKBOX) ? item.asCheckboxItem() : item.asMultipleChoiceItem();
        const correctnessFeedbackObject = createCorrectnessFeedbackObject(
          typedItem
        );
        return {
          isRequired: typedItem.isRequired && typedItem.isRequired(),
          hasOtherOption: typedItem.hasOtherOption && typedItem.hasOtherOption(),
          choices: typedItem.getChoices && typedItem.getChoices().map(choiceToJson),
          ...(correctnessFeedbackObject ? correctnessFeedbackObject : {}),
          ...typedItemObject,
        } as CheckboxItemObject | MultipleChoiceItemObject;
      })();

    case FormApp.ItemType.LIST:
      return (() => {
        const typedItem = item.asListItem();
        const correctnessFeedbackObject = createCorrectnessFeedbackObject(
          typedItem
        );
        return {
          isRequired: typedItem.isRequired(),
          choices: typedItem.getChoices().map(choiceToJson),
          ...(correctnessFeedbackObject ? correctnessFeedbackObject : {}),
          ...typedItemObject,
        } as ListItemObject;
      })();

    case FormApp.ItemType.DATE:
    case FormApp.ItemType.DATETIME:
      return (() => {
        const typedItem = (item.getType() === FormApp.ItemType.DATE)? item.asDateItem() : item.asDateTimeItem();
        const quizFeedback = typedItem.getGeneralFeedback();
        const generalFeedbackObject = createGeneralFeedbackObject(quizFeedback, typedItem.getPoints());
        return {
          isRequired: typedItem.isRequired(),
          includesYear: typedItem.includesYear(),
          ...(generalFeedbackObject ? generalFeedbackObject : {}),
          ...typedItemObject,
        } as DateItemObject | DateTimeItemObject;
      })();

    case FormApp.ItemType.TIME:
      return (() => {
        const typedItem = item.asTimeItem();
        const quizFeedback = typedItem.getGeneralFeedback();
        const generalFeedbackObject = createGeneralFeedbackObject(quizFeedback, typedItem.getPoints());
        return {
          isRequired: typedItem.isRequired(),
          ...(generalFeedbackObject ? generalFeedbackObject : {}),
          ...typedItemObject,
        } as TimeItemObject;
      })();

    case FormApp.ItemType.DURATION:
      return (() => {
        const typedItem = item.asDurationItem();
        const quizFeedback = typedItem.getGeneralFeedback();
        const generalFeedbackObject = createGeneralFeedbackObject(quizFeedback, typedItem.getPoints());
        return {
          isRequired: typedItem.isRequired(),
          ...(generalFeedbackObject ? generalFeedbackObject : {}),
          ...typedItemObject,
        } as DurationItemObject;
      })();

    case FormApp.ItemType.CHECKBOX_GRID:
      return (() => {
        const typedItem = item.asCheckboxGridItem();
        // TODO: The CheckboxGridItem type of Forms API doesn't have getGeneralFeedback and getPoints method for now.
        //const quizFeedback = typedItem.getGeneralFeedback();
        //const generalFeedbackObject = createGeneralFeedbackObject(quizFeedback, typedItem.getPoints());
        return {
          isRequired: typedItem.isRequired(),
          rows: typedItem.getRows(),
          columns: typedItem.getColumns(),
          // ...(generalFeedbackObject ? generalFeedbackObject : {}),
          ...typedItemObject,
        } as CheckboxGridItemObject | GridItemObject;
      })();
    case FormApp.ItemType.GRID:
      return (() => {
        const typedItem = item.asGridItem();
        // TODO: The GridItem type of Forms API doesn't have getGeneralFeedback and getPoints method for now.
        // const quizFeedback = typedItem.getGeneralFeedback();
        // const generalFeedbackObject = createGeneralFeedbackObject(quizFeedback, typedItem.getPoints());
        return {
          isRequired: typedItem.isRequired(),
          rows: typedItem.getRows(),
          columns: typedItem.getColumns(),
          // ...(generalFeedbackObject ? generalFeedbackObject : {}),
          ...typedItemObject,
        } as CheckboxGridItemObject | GridItemObject;
      })();

    case FormApp.ItemType.PARAGRAPH_TEXT:
      return (() => {
        const typedItem = item.asParagraphTextItem();
        const quizFeedback = typedItem.getGeneralFeedback();
        const generalFeedbackObject = createGeneralFeedbackObject(quizFeedback, typedItem.getPoints());
        return {
          isRequired: typedItem.isRequired(),
          ...(generalFeedbackObject ? generalFeedbackObject : {}),
          ...typedItemObject,
          // validation: undefined
        } as ParagraphTextItemObject;
      })();

    case FormApp.ItemType.SCALE:
      return (() => {
        const typedItem = item.asScaleItem();
        const quizFeedback = typedItem.getGeneralFeedback();
        const generalFeedbackObject = createGeneralFeedbackObject(quizFeedback, typedItem.getPoints());
        return {
          isRequired: typedItem.isRequired(),
          leftLabel: typedItem.getLeftLabel(),
          lowerBound: typedItem.getLowerBound(),
          rightLabel: typedItem.getRightLabel(),
          upperBound: typedItem.getUpperBound(),
          ...(generalFeedbackObject ? generalFeedbackObject : {}),
          ...typedItemObject,
        } as ScaleItemObject;
      })();

    case FormApp.ItemType.TEXT:
      return (() => {
        const typedItem = item.asTextItem();
        const quizFeedback = typedItem.getGeneralFeedback();
        const generalFeedbackObject = createGeneralFeedbackObject(quizFeedback, typedItem.getPoints());
        return {
          isRequired: typedItem.isRequired(),
          // validation: undefined
          ...(generalFeedbackObject ? generalFeedbackObject : {}),
          ...typedItemObject,
        } as TextItemObject;
      })();

    case FormApp.ItemType.IMAGE:
      return (() => {
        const typedItem = item.asImageItem();
        return {
          image: blobToJson(typedItem.getImage()),
          alignment: getAlignmentString(typedItem.getAlignment()),
          width: typedItem.getWidth(),
          ...typedItemObject,
        } as ImageItemObject;
      })();

    case FormApp.ItemType.VIDEO:
      return (() => {
        const typedItem = item.asVideoItem();
        return {
          // videoUrl: typedItem.getVideoUrl(),
          alignment: getAlignmentString(typedItem.getAlignment()),
          width: typedItem.getWidth(),
          ...typedItemObject,
        } as VideoItemObject;
      })();

    case FormApp.ItemType.PAGE_BREAK:
      return (() => {
        const typedItem = item.asPageBreakItem();
        const pageNavigationType = item.asPageBreakItem().getPageNavigationType();
        if (typedItem.getGoToPage && typedItem.getGoToPage()) {
          const goToPageTitle = typedItem.getGoToPage().getTitle();
          return {
            pageNavigationType: getPageNavigationTypeString(
              pageNavigationType
            ),
            goToPageTitle,
            ...typedItemObject,
          } as PageBreakItemObject;
        } else {
          return {
            pageNavigationType: getPageNavigationTypeString(
              pageNavigationType
            ),
            ...typedItemObject,
          } as PageBreakItemObject;
        }
      })();

    case FormApp.ItemType.SECTION_HEADER:
      return { ...createItemObject(item) } as SectionHeaderItemObject;

    default:
      throw new Error("InvalidItem:" + item.getType());
  }
}
