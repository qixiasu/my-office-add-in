/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, Excel, window */

var { fixGarbledText } = require("../utils/encoding-utils");

// Expose to global scope for inline script in commands.html
window.fixGarbledText = fixGarbledText;
