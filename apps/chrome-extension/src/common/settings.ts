import { pick } from 'es-toolkit'
import { defaultsDeep } from 'es-toolkit/compat'
import { supported } from 'browser-fs-access'
import { EventName, portImpl } from './message'

export enum SettingKey {
  Locale = 'general.locale',
  Theme = 'general.theme',
  DownloadMethod = 'download.method',
  Table = 'general.table',
  Grid = 'general.grid',
  TextHighlight = 'general.text_highlight',
  DownloadFileWithUniqueName = 'download.file_with_unique_name',
  HtmlImageInline = 'html.image_inline',
  HtmlMaxInlineSizeKb = 'html.max_inline_size_kb',
  HtmlImageCompressEnabled = 'html.image_compress_enabled',
  HtmlImageCompressQuality = 'html.image_compress_quality',
  HtmlIncludeStyles = 'html.include_styles',
  HtmlPrintFriendly = 'html.print_friendly',
}

export enum Theme {
  Light = 'light',
  Dark = 'dark',
  System = 'system',
}

export enum DownloadMethod {
  Direct = 'direct',
  ShowSaveFilePicker = 'showSaveFilePicker',
}

export enum Table {
  Filtered = 'filtered',
  NonPhrasingContentToHTML = 'nonPhrasingContentToHTML',
  ToHTML = 'toHTML',
}

export enum Grid {
  Flatten = 'flatten',
  ToTable = 'toTable',
  ToHTML = 'toHTML',
}

export interface Settings {
  [SettingKey.Locale]: string
  [SettingKey.Theme]: (typeof Theme)[keyof typeof Theme]
  [SettingKey.DownloadMethod]: (typeof DownloadMethod)[keyof typeof DownloadMethod]
  [SettingKey.Table]: (typeof Table)[keyof typeof Table]
  [SettingKey.Grid]: (typeof Grid)[keyof typeof Grid]
  [SettingKey.TextHighlight]: boolean
  [SettingKey.DownloadFileWithUniqueName]: boolean
  [SettingKey.HtmlImageInline]: boolean
  [SettingKey.HtmlMaxInlineSizeKb]: number
  [SettingKey.HtmlImageCompressEnabled]: boolean
  [SettingKey.HtmlImageCompressQuality]: number
  [SettingKey.HtmlIncludeStyles]: boolean
  [SettingKey.HtmlPrintFriendly]: boolean
}

export const fallbackSettings: Settings = {
  [SettingKey.Locale]: 'en-US',
  [SettingKey.Theme]: Theme.System,
  [SettingKey.DownloadMethod]: supported
    ? DownloadMethod.ShowSaveFilePicker
    : DownloadMethod.Direct,
  [SettingKey.Table]: Table.NonPhrasingContentToHTML,
  [SettingKey.Grid]: Grid.Flatten,
  [SettingKey.TextHighlight]: true,
  [SettingKey.DownloadFileWithUniqueName]: false,
  [SettingKey.HtmlImageInline]: true,
  [SettingKey.HtmlMaxInlineSizeKb]: 10240,
  [SettingKey.HtmlImageCompressEnabled]: true,
  [SettingKey.HtmlImageCompressQuality]: 0.6,
  [SettingKey.HtmlIncludeStyles]: true,
  [SettingKey.HtmlPrintFriendly]: false,
}

export const getSettings = async <Key extends keyof Settings>(
  keys: Key[],
): Promise<Pick<Settings, Key>> => {
  try {
    const settings = await portImpl.sender.sendAsync(
      EventName.GetSettings,
      keys,
    )
    return pick(defaultsDeep(settings, fallbackSettings), keys)
  } catch (error) {
    console.error(error)

    return pick(fallbackSettings, keys)
  }
}
