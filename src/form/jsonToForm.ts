import {
  CheckboxItemObject,
  ChoiceObject,
  CorrectnessFeedbackObject,
  CorrectnessFeedbackItem,
  FormSource,
  HasOtherOptionObject,
  ShowOtherOptionItem,
  IncludesYearItem,
  IncludesYearObject,
  PageBreakItemObject,
  QuizItem,
  QuizItemObject,
  MultipleChoiceItemObject,
  GridItemObject,
  ListItemObject,
  CheckboxGridItemObject,
  DateItemObject,
  DateTimeItemObject,
  TextItemObject,
  TimeItemObject,
  ParagraphTextItemObject,
  DurationItemObject,
  SectionHeaderItemObject,
  VideoItemObject,
  ImageItemObject,
  ScaleItemObject,
  ItemObject,
  getAlignment,
  QuizFeedbackObject,
} from "./types";
type PageBreakItem = GoogleAppsScript.Forms.PageBreakItem;
type MultipleChoiceItem = GoogleAppsScript.Forms.MultipleChoiceItem;
type CheckboxItem = GoogleAppsScript.Forms.CheckboxItem;
type GridItem = GoogleAppsScript.Forms.GridItem;
type CheckboxGridItem = GoogleAppsScript.Forms.CheckboxGridItem;
type ListItem = GoogleAppsScript.Forms.ListItem;
type Form = GoogleAppsScript.Forms.Form;
type Item = GoogleAppsScript.Forms.Item;
type User = GoogleAppsScript.Base.User;

type ItemHandlers = {
  //{[key: string]: (itemObject: ItemObject, forms: Form)=>void};
  multipleChoice: (itemObject: MultipleChoiceItemObject, form: Form) => void;
  checkbox: (itemObject: CheckboxItemObject, form: Form) => void;
  list: (itemObject: ListItemObject, form: Form) => void;
  checkboxGrid: (itemObject: CheckboxGridItemObject, form: Form) => void;
  grid: (itemObject: GridItemObject, form: Form) => void;
  // item: (itemObject: ItemObject, forms: Form) => void;
  date: (itemObject: DateItemObject, form: Form) => void;
  datetime: (itemObject: DateTimeItemObject, form: Form) => void;
  text: (itemObject: TextItemObject, form: Form) => void;
  time: (itemObject: TimeItemObject, form: Form) => void;
  paragraphText: (itemObject: ParagraphTextItemObject, form: Form) => void;
  duration: (itemObject: DurationItemObject, form: Form) => void;
  sectionHeader: (itemObject: SectionHeaderItemObject, form: Form) => void;
  video: (itemObject: VideoItemObject, form: Form) => void;
  image: (itemObject: ImageItemObject, form: Form) => void;
  scale: (itemObject: ScaleItemObject, form: Form) => void;
  pageBreak: (itemObject: PageBreakItemObject, form: Form) => void;
};

function booleanValue(value: boolean | string | number) {
  if (typeof value === "boolean") {
    return value;
  } else if (typeof value === "string") {
    return !(value.toLowerCase() === "false" || value === "0" || value === "");
  } else if (typeof value === "number") {
    return 0 < value;
  }
  return false;
}

function callWithBooleanValue(
  value: string | boolean | number,
  callback: (value: boolean) => void
) {
  const b = booleanValue(value);
  if (b) {
    callback(b);
  }
}

function isNotNullValue(value: boolean | string | number) {
  return value !== undefined && value !== null && value !== "";
}

function callWithNotNullValue(
  value: string,
  callback: (value: string) => void
) {
  if (isNotNullValue(value)) {
    callback(value);
  }
}

export function jsonToForm(json: FormSource, form: Form | null): Form {
  if (form === null) {
    if (json.metadata.id) {
      const form = FormApp.openById(json.metadata.id);
      if (form) {
        const numItems = form.getItems().length;
        for (let index = numItems - 1; 0 <= index; index--) {
          form.deleteItem(index);
        }
        form.getEditors().forEach((editor: User) => {
          form.removeEditor(editor);
        });
      }
    } else {
      form = FormApp.create(json.metadata.title);
    }
  }
  if (!form) {
    throw new Error("invalid forms title");
  }

  json.metadata.quiz && form.setIsQuiz(json.metadata.quiz);
  json.metadata.allowResponseEdits &&
    form.setAllowResponseEdits(json.metadata.allowResponseEdits);
  json.metadata.collectEmail &&
    form.setCollectEmail(json.metadata.collectEmail);
  json.metadata.description && form.setDescription(json.metadata.description);
  json.metadata.acceptingResponses &&
    form.setAcceptingResponses(json.metadata.acceptingResponses);
  json.metadata.publishingSummary &&
    form.setPublishingSummary(json.metadata.publishingSummary);
  json.metadata.confirmationMessage &&
    form.setConfirmationMessage(json.metadata.confirmationMessage);
  json.metadata.customClosedFormMessage &&
    form.setCustomClosedFormMessage(json.metadata.customClosedFormMessage);
  json.metadata.limitOneResponsePerUser &&
    form.setLimitOneResponsePerUser(json.metadata.limitOneResponsePerUser);
  json.metadata.progressBar && form.setProgressBar(json.metadata.progressBar);
  json.metadata.editors && form.addEditors(json.metadata.editors);
  json.metadata.shuffleQuestions &&
    form.setShuffleQuestions(json.metadata.shuffleQuestions);
  json.metadata.linkToRespondAgain &&
    form.setShowLinkToRespondAgain(json.metadata.linkToRespondAgain);
  form.setTitle(json.metadata.title);
  if (
    json.metadata.destinationType ==
      FormApp.DestinationType.SPREADSHEET.toString() &&
    json.metadata.destinationId
  ) {
    form.setDestination(
      FormApp.DestinationType.SPREADSHEET,
      json.metadata.destinationId
    );
  }

  const pageBreakItems = new Map<string, PageBreakItem>();
  json.items
    .filter((item) => item.type === "pageBreakItem")
    .forEach((item) => {
      const pageBreakItemObject = item as PageBreakItemObject;
      if (!form) {
        throw new Error("invalid forms title");
      }
      const title = pageBreakItemObject.title;
      const pageBreakItem = form.addPageBreakItem().setTitle(title);
      pageBreakItems.set(title, pageBreakItem);
      return pageBreakItem;
    });

  json.items.forEach((itemObject) => {
    if (!form) {
      throw new Error("forms is null");
    }
    handleItem(itemObject, pageBreakItems, form);
  });

  return form;
}

function handleItem(
  itemObject: ItemObject,
  pageBreakItems: Map<string, PageBreakItem>,
  form: Form
) {
  function createQuizFeedback(feedbackObject: QuizFeedbackObject) {
    const { text, links } = feedbackObject;
    const feedbackBuilder = FormApp.createFeedback();
    if (text) {
      feedbackBuilder.setText(text);
    }
    links.forEach((link) => {
      const { url, displayText } = link;
      if (displayText) {
        feedbackBuilder.addLink(url, displayText);
      } else {
        feedbackBuilder.addLink(url);
      }
    });
    return feedbackBuilder.build();
  }

  const itemModifiers = {
    itemProperties: function (itemObject: ItemObject, item: Item) {
      callWithNotNullValue(itemObject.title, function (value: string) {
        item.setTitle(value);
      });
      callWithNotNullValue(itemObject.helpText, function (value: string) {
        item.setHelpText(value);
      });
    },
    questionProperties: (itemObject: QuizItemObject, item: QuizItem) => {
      callWithBooleanValue(itemObject.isRequired, function (value: boolean) {
        item.setRequired(value);
      });
    },
    choices: function (
      choiceObjectList: ChoiceObject[],
      item: MultipleChoiceItem | CheckboxItem | ListItem
    ) {
      const choiceObjects = choiceObjectList.map(function (
        choiceObject: ChoiceObject
      ) {
        if (!form) {
          throw new Error("invalid forms title");
        }
        if (!choiceObject.gotoPageTitle) {
          throw new Error("invalid value of gotoPageTitle");
        }

        const goToPage = pageBreakItems.get(choiceObject.gotoPageTitle);
        const pageNavigationType =
          choiceObject.pageNavigationType === "CONTINUE"
            ? FormApp.PageNavigationType.CONTINUE
            : choiceObject.pageNavigationType === "GO_TO_PAGE"
            ? FormApp.PageNavigationType.GO_TO_PAGE
            : choiceObject.pageNavigationType === "RESTART"
            ? FormApp.PageNavigationType.RESTART
            : FormApp.PageNavigationType.SUBMIT;

        if (!goToPage && !choiceObject.isCorrectAnswer) {
          return item.createChoice(choiceObject.value);
        } else if (!goToPage && choiceObject.isCorrectAnswer) {
          return item.createChoice(
            choiceObject.value,
            choiceObject.isCorrectAnswer
          );
        } else if (goToPage) {
          return item.createChoice(choiceObject.value, pageNavigationType);
        }
        throw new Error("invalid choiceObject:" + choiceObject);
      });
      item.setChoices(choiceObjects);
    },

    showOtherOption: function (
      itemObject: HasOtherOptionObject,
      item: ShowOtherOptionItem
    ) {
      callWithBooleanValue(
        itemObject.hasOtherOption,
        function (value: boolean) {
          item.showOtherOption(value);
        }
      );
    },
    quizProperties: function (
      itemObject: CorrectnessFeedbackObject,
      item: CorrectnessFeedbackItem
    ) {
      item.setPoints(itemObject.points);
      itemObject.feedbackForCorrect &&
        item.setFeedbackForCorrect(
          createQuizFeedback(itemObject.feedbackForCorrect)
        );
      itemObject.feedbackForIncorrect &&
        item.setFeedbackForIncorrect(
          createQuizFeedback(itemObject.feedbackForIncorrect)
        );
    },
    includesYear: function (
      itemObject: IncludesYearObject,
      item: IncludesYearItem
    ) {
      callWithBooleanValue(itemObject.includesYear, function (value: boolean) {
        item.setIncludesYear(value);
      });
    },
  };

  function multipleChoiceHandler(
    itemObject: MultipleChoiceItemObject,
    form: Form
  ) {
    const item = form.addMultipleChoiceItem();
    itemModifiers.choices(itemObject.choices, item);
    itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
    itemModifiers.questionProperties(itemObject, (item as unknown) as QuizItem);
    itemModifiers.showOtherOption(
      itemObject,
      (item as unknown) as ShowOtherOptionItem
    );
    if (form.isQuiz()) {
      itemModifiers.quizProperties(
        itemObject,
        (item as unknown) as CorrectnessFeedbackItem
      );
    }
  }

  function gridHandler(
    itemObject: GridItemObject | CheckboxGridItemObject,
    item: GridItem | CheckboxGridItem
  ) {
    item.setRows(itemObject.rows).setColumns(itemObject.columns);
    itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
    itemModifiers.questionProperties(itemObject, (item as unknown) as QuizItem);
  }

  const itemHandlers: ItemHandlers = {
    multipleChoice: multipleChoiceHandler,

    checkbox: function (itemObject: CheckboxItemObject, form: Form) {
      const item = form.addCheckboxItem();
      itemModifiers.choices(itemObject.choices, item);
      itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
      itemModifiers.questionProperties(
        itemObject,
        (item as unknown) as QuizItem
      );
      itemModifiers.showOtherOption(
        itemObject,
        (item as unknown) as ShowOtherOptionItem
      );
      if (form.isQuiz()) {
        itemModifiers.quizProperties(
          itemObject,
          (item as unknown) as CorrectnessFeedbackItem
        );
      }
    },

    list: function (itemObject: ListItemObject, form: Form) {
      const item = form.addListItem();
      itemModifiers.choices(itemObject.choices, item);
      itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
      itemModifiers.questionProperties(
        itemObject,
        (item as unknown) as QuizItem
      );
    },

    checkboxGrid: function (itemObject: CheckboxGridItemObject, form: Form) {
      const item = form.addCheckboxGridItem();
      gridHandler(itemObject, item);
    },

    grid: function (itemObject: GridItemObject, form: Form) {
      const item = form.addGridItem();
      gridHandler(itemObject, item);
    },

    time: function (itemObject: TimeItemObject, form: Form) {
      const item = form.addTimeItem();
      itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
      itemModifiers.questionProperties(
        itemObject,
        (item as unknown) as QuizItem
      );
    },

    date: function (itemObject: DateItemObject, form: Form) {
      const item = form.addDateItem();
      itemModifiers.includesYear(
        itemObject,
        (item as unknown) as IncludesYearItem
      );
      itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
      itemModifiers.questionProperties(
        itemObject,
        (item as unknown) as QuizItem
      );
    },

    datetime: function (itemObject: DateTimeItemObject, form: Form) {
      const item = form.addDateTimeItem();
      itemModifiers.includesYear(
        itemObject,
        (item as unknown) as IncludesYearItem
      );
      itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
      itemModifiers.questionProperties(
        itemObject,
        (item as unknown) as QuizItem
      );
    },

    text: function (itemObject: TextItemObject, form: Form) {
      const item = form.addTextItem();
      itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
      itemModifiers.questionProperties(
        itemObject,
        (item as unknown) as QuizItem
      );
    },

    paragraphText: function (itemObject: ParagraphTextItemObject, form: Form) {
      const item = form.addParagraphTextItem();
      itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
      itemModifiers.questionProperties(
        itemObject,
        (item as unknown) as QuizItem
      );
    },

    duration: function (itemObject: DurationItemObject, form: Form) {
      const item = form.addDurationItem();
      itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
      itemModifiers.questionProperties(
        itemObject,
        (item as unknown) as QuizItem
      );
    },

    scale: function (itemObject: ScaleItemObject, form: Form) {
      const item = form.addScaleItem();
      callWithNotNullValue(itemObject.leftLabel, function (left: string) {
        callWithNotNullValue(itemObject.rightLabel, function (right: string) {
          item.setLabels(left, right);
        });
      });
      item.setBounds(itemObject.lowerBound, itemObject.upperBound);
      itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
      itemModifiers.questionProperties(
        itemObject,
        (item as unknown) as QuizItem
      );
    },

    sectionHeader: function (itemObject: SectionHeaderItemObject, form: Form) {
      const item = form.addSectionHeaderItem();
      itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
    },

    video: function (itemObject: VideoItemObject, form: Form) {
      const item = form.addVideoItem();
      const videoUrl = itemObject.videoUrl;
      item.setVideoUrl(videoUrl);
      item.setWidth(itemObject.width);
      callWithNotNullValue(itemObject.alignment, function (value: string) {
        const alignment = getAlignment(value);
        if (alignment) {
          item.setAlignment(alignment);
        }
      });
      itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
    },

    image: function (itemObject: ImageItemObject, form: Form) {
      const item = form.addImageItem();
      const image = UrlFetchApp.fetch(itemObject.url);
      item.setImage(image);
      item.setWidth(itemObject.width);
      const alignment = getAlignment(itemObject.alignment);
      item.setAlignment(alignment);
      itemModifiers.itemProperties(itemObject, (item as unknown) as Item);
    },

    pageBreak: function (itemObject: PageBreakItemObject, form: Form) {
      const title = itemObject.title;
      const pageNavigationType = itemObject.pageNavigationType;

      const pageBreakItem = pageBreakItems.get(title);
      if (!pageBreakItem) {
        return;
      }
      const lastItemIndex = form.getItems().length - 1;
      form.moveItem(pageBreakItem.getIndex(), lastItemIndex);

      if (pageNavigationType === "CONTINUE") {
        pageBreakItem.setGoToPage(FormApp.PageNavigationType.CONTINUE);
      } else if (pageNavigationType === "RESTART") {
        pageBreakItem.setGoToPage(FormApp.PageNavigationType.RESTART);
      } else if (pageNavigationType === "SUBMIT") {
        pageBreakItem.setGoToPage(FormApp.PageNavigationType.SUBMIT);
      } else if (itemObject.gotoPageTitle) {
        const gotoPage = pageBreakItems.get(itemObject.gotoPageTitle);
        if (gotoPage) {
          pageBreakItem.setGoToPage(gotoPage);
        }
      }
      itemModifiers.itemProperties(
        itemObject,
        (pageBreakItem as unknown) as Item
      );
    },
  };

  switch (itemObject.type) {
    case "multipleChoice":
      return itemHandlers.multipleChoice(
        itemObject as MultipleChoiceItemObject,
        form
      );
    case "checkbox":
      return itemHandlers.checkbox(itemObject as CheckboxItemObject, form);
    case "list":
      return itemHandlers.list(itemObject as ListItemObject, form);
    case "checkboxGrid":
      return itemHandlers.checkboxGrid(
        itemObject as CheckboxGridItemObject,
        form
      );
    case "grid":
      return itemHandlers.grid(itemObject as GridItemObject, form);
    case "time":
      return itemHandlers.time(itemObject as TimeItemObject, form);
    case "date":
      return itemHandlers.date(itemObject as DateItemObject, form);
    case "datetime":
      return itemHandlers.datetime(itemObject as DateTimeItemObject, form);
    case "text":
      return itemHandlers.text(itemObject as TextItemObject, form);
    case "paragraphText":
      return itemHandlers.paragraphText(
        itemObject as ParagraphTextItemObject,
        form
      );
    case "duration":
      return itemHandlers.duration(itemObject as DurationItemObject, form);
    case "scale":
      return itemHandlers.scale(itemObject as ScaleItemObject, form);
    case "sectionHeader":
      return itemHandlers.sectionHeader(
        itemObject as SectionHeaderItemObject,
        form
      );
    case "video":
      return itemHandlers.video(itemObject as VideoItemObject, form);
    case "image":
      return itemHandlers.image(itemObject as ImageItemObject, form);
    case "pageBreak":
      return itemHandlers.pageBreak(itemObject, form);
    default:
      throw new Error("Invalid type: " + itemObject.type);
  }
}
