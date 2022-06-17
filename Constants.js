// Copyright 2018 Google Inc. All Rights Reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/**
 * @fileoverview Constants used throughout the addon are declared here.
 */

/**
 * Service name to use for oauth service.
 * @const {string}
 */
var SERVICENAME = 'dfareporting';

/**
 * Name of the addon.
 * @const {string}
 */
var ADDON_TITLE = 'CM Report - Google Sheets Addon';

/**
 * Scopes used by the addon.
 * @const {string}
 */
var SCOPE = 'https://www.googleapis.com/auth/dfareporting';

/**
 * URI to request oauth token.
 * @const {string}
 */
var REQUEST_TOKEN_URL = 'https://www.google.com/accounts/OAuthGetRequestToken';

/**
 * URI to access oauth token.
 * @const {string}
 */
var TOKEN_ACCESS_URL = 'https://accounts.google.com/o/oauth2/token';

/**
 * Oauth authorization URL.
 * @const {string}
 */
var AUTHORIZATION_URL = 'https://accounts.google.com/o/oauth2/auth';

/**
 * Google developer console client id.
 * @const {string}
 */
var CLIENT_ID = 'Insert client id here';

/**
 * Google developer console client secret.
 * @const {string}
 */
var CLIENT_SECRET = 'Insert client secret here';

/**
 * Text mapping for daily frequency values.
 * @const {Array<Object>}
 */
var DAILY_FREQUENCY = [
  {'value' : 0, 'text' : 'Midnight to 1am'},
  {'value' : 1, 'text' : '1am to 2am'},
  {'value' : 2, 'text' : '2am to 3am'},
  {'value' : 3, 'text' : '3am to 4am'},
  {'value' : 4, 'text' : '4am to 5am'},
  {'value' : 5, 'text' : '5am to 6am'},
  {'value' : 6, 'text' : '6am to 7am'},
  {'value' : 7, 'text' : '7am to 8am'},
  {'value' : 8, 'text' : '8am to 9am'},
  {'value' : 9, 'text' : '9am to 10am'},
  {'value' : 10, 'text' : '10am to 11am'},
  {'value' : 11, 'text' : '11am to noon'},
  {'value' : 12, 'text' : 'noon to 1pm'},
  {'value' : 13, 'text' : '1pm to 2pm'},
  {'value' : 14, 'text' : '2pm to 3pm'},
  {'value' : 15, 'text' : '3pm to 4pm'},
  {'value' : 16, 'text' : '4pm to 5pm'},
  {'value' : 17, 'text' : '5pm to 6pm'},
  {'value' : 18, 'text' : '6pm to 7pm'},
  {'value' : 19, 'text' : '7pm to 8pm'},
  {'value' : 20, 'text' : '8pm to 9pm'},
  {'value' : 21, 'text' : '9pm to 10pm'},
  {'value' : 22, 'text' : '10pm to 11pm'},
  {'value' : 23, 'text' : '11pm to midnight'}
];

/**
 * Title of the dialog.
 * @const {string}
 */
var DIALOG_TITLE = 'CM Report Selection';

/**
 * Property identifier for frequency.
 * @const {string}
 */
var PROP_DCM_SCHEDULE_FREQUENCY = 'DCM_Schedule_Frequency';

/**
 * Property identifier for time.
 * @const {string}
 */
var PROP_DCM_SCHEDULE_TIME = 'DCM_Schedule_Time';

/**
 * Property identifier for time2.
 * @const {string}
 */
var PROP_DCM_SCHEDULE_TIME2 = 'DCM_Schedule_Time2';

/**
 * Property identifier for owner.
 * @const {string}
 */
var PROP_DCM_TRIGGER_CREATED_BY = 'DCM_Trigger_Created_By';

/**
 * Property identifier for trigger id.
 * @const {string}
 */
var PROP_DCM_TRIGGER_ID = 'DCM_Trigger_ID';

/**
 * Suffix for network id.
 * @const {string}
 */
var SUFFIX_NETWORK_ID = '_NETWORK_ID';

/**
 * Suffix for report name.
 * @const {string}
 */
var SUFFIX_REPORT_NAME = '_REPORT_NAME';

/**
 * Suffix for report setup user.
 * @const {string}
 */
var SUFFIX_DCM_REPORT_SETUP_USER = '_DCM_REPORT_SETUP_USER';

/**
 * Suffix for profile id.
 * @const {string}
 */
var SUFFIX_PROFILE_ID = '_PROFILE_ID';

/**
 * Suffix for report id.
 * @const {string}
 */
var SUFFIX_REPORT_ID = '_REPORT_ID';

/**
 * Suffix for bucket name.
 * @const {string}
 */
var SUFFIX_BUCKET_NAME = '_BUCKET_NAME';

/**
 * Suffix for dbm report updated date.
 * @const {string}
 */
var SUFFIX_DBM_REPORT_UPDATED_DATE = '_DBM_REPORT_UPDATED_DATE';

/**
 * Suffix for webquery url.
 * @const {string}
 */
var SUFFIX_WEBQUERY_URL = '_WEBQUERY_URL';

/**
 * Suffix for last sync.
 * @const {string}
 */
var SUFFIX_LAST_SYNC = '_LAST_SYNC';

/**
 * Trigger name for offline report sync.
 * @const {string}
 */
var TRIGGER_DCM_OFFLINE = 'DCM_offlineReportSync';
