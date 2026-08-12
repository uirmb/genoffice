import { aiStrings } from './strings-ai'
import { appStrings } from './strings-app'
import { dialogStrings } from './strings-dialogs'
import { fileStrings } from './strings-file'

export const strings = {
  zh: { ...appStrings.zh, ...dialogStrings.zh, ...aiStrings.zh, ...fileStrings.zh },
  en: { ...appStrings.en, ...dialogStrings.en, ...aiStrings.en, ...fileStrings.en },
  ja: { ...appStrings.ja, ...dialogStrings.ja, ...aiStrings.ja, ...fileStrings.ja },
  ko: { ...appStrings.ko, ...dialogStrings.ko, ...aiStrings.ko, ...fileStrings.ko },
  fr: { ...appStrings.fr, ...dialogStrings.fr, ...aiStrings.fr, ...fileStrings.fr },
  de: { ...appStrings.de, ...dialogStrings.de, ...aiStrings.de, ...fileStrings.de },
  es: { ...appStrings.es, ...dialogStrings.es, ...aiStrings.es, ...fileStrings.es },
  th: { ...appStrings.th, ...dialogStrings.th, ...aiStrings.th, ...fileStrings.th },
  id: { ...appStrings.id, ...dialogStrings.id, ...aiStrings.id, ...fileStrings.id },
  ru: { ...appStrings.ru, ...dialogStrings.ru, ...aiStrings.ru, ...fileStrings.ru },
  ar: { ...appStrings.ar, ...dialogStrings.ar, ...aiStrings.ar, ...fileStrings.ar },
  pt: { ...appStrings.pt, ...dialogStrings.pt, ...aiStrings.pt, ...fileStrings.pt },
  it: { ...appStrings.it, ...dialogStrings.it, ...aiStrings.it, ...fileStrings.it },
  pl: { ...appStrings.pl, ...dialogStrings.pl, ...aiStrings.pl, ...fileStrings.pl },
  nl: { ...appStrings.nl, ...dialogStrings.nl, ...aiStrings.nl, ...fileStrings.nl },
  ms: { ...appStrings.ms, ...dialogStrings.ms, ...aiStrings.ms, ...fileStrings.ms },
  he: { ...appStrings.he, ...dialogStrings.he, ...aiStrings.he, ...fileStrings.he },
  hi: { ...appStrings.hi, ...dialogStrings.hi, ...aiStrings.hi, ...fileStrings.hi },
  'zh-TW': {
    ...appStrings['zh-TW'],
    ...dialogStrings['zh-TW'],
    ...aiStrings['zh-TW'],
    ...fileStrings['zh-TW'],
  },
}
