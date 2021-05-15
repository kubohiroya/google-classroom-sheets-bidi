import "./functions/picker";
import "./functions/classroom";
import "./functions/form";
import "./functions/onOpen";

import { onOpen } from "./functions/onOpen";
import {getOAuthToken, pickerHandler} from './functions/picker/picker';
global.onOpen = onOpen;
global.getOAuthToken = getOAuthToken;
global.pickerHandler = pickerHandler;
