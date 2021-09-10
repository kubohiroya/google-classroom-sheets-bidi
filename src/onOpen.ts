import {exportFormSubMenu, importFormSubMenu} from './formSubMenu';
import {exportClassroomSubMenu, importClassroomSubMenu} from './classroomSubMenu';

export function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("ClassroomSync")
    .addSubMenu(importClassroomSubMenu(ui))
    .addSubMenu(exportClassroomSubMenu(ui))
    .addSeparator()
    .addSubMenu(importFormSubMenu(ui))
    .addSubMenu(exportFormSubMenu(ui))
    .addToUi();
}


