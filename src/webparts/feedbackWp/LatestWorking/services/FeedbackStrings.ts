// Moved all button/message text here so FeedbackWp.tsx doesn't become
// a huge file just because of translations. Language key is same as the
// dropdown property in FeedbackWpWebPart.ts - whatever page owner selects

export interface IFeedbackStrings {
  YesButtonLabel: string;
  NoButtonLabel: string;
  CommentPlaceholder: string;
  SubmitButtonLabel: string;
  SubmittingButtonLabel: string;
  CancelButtonLabel: string;
  ThankYouTitle: string;
  ThankYouSubtitle: string;
}

const STRINGS: { [languageKey: string]: IFeedbackStrings } = {
  en: {
    YesButtonLabel: 'Yes',
    NoButtonLabel: 'No',
    CommentPlaceholder: '(Optional) Please provide feedback so we can continuously improve 3M Go.',
    SubmitButtonLabel: 'Submit',
    SubmittingButtonLabel: 'Submitting…',
    CancelButtonLabel: 'Cancel',
    ThankYouTitle: 'Thank you for your Feedback!',
    ThankYouSubtitle: 'Your input helps us continuously improve 3M Go.'
  },
  es: {
    YesButtonLabel: 'Sí',
    NoButtonLabel: 'No',
    CommentPlaceholder: '(Opcional) Por favor, comparta sus comentarios para que podamos mejorar continuamente 3M Go.',
    SubmitButtonLabel: 'Enviar',
    SubmittingButtonLabel: 'Enviando…',
    CancelButtonLabel: 'Cancelar',
    ThankYouTitle: '¡Gracias por sus comentarios!',
    ThankYouSubtitle: 'Su opinión nos ayuda a mejorar continuamente 3M Go.'
  },
  de: {
    YesButtonLabel: 'Ja',
    NoButtonLabel: 'Nein',
    CommentPlaceholder: '(Optional) Bitte geben Sie Feedback, damit wir 3M Go kontinuierlich verbessern können.',
    SubmitButtonLabel: 'Absenden',
    SubmittingButtonLabel: 'Wird gesendet…',
    CancelButtonLabel: 'Abbrechen',
    ThankYouTitle: 'Vielen Dank für Ihr Feedback!',
    ThankYouSubtitle: 'Ihre Rückmeldung hilft uns, 3M Go kontinuierlich zu verbessern.'
  },
  fr: {
    YesButtonLabel: 'Oui',
    NoButtonLabel: 'Non',
    CommentPlaceholder: '(Facultatif) Merci de nous faire part de vos commentaires afin que nous puissions améliorer continuellement 3M Go.',
    SubmitButtonLabel: 'Envoyer',
    SubmittingButtonLabel: 'Envoi en cours…',
    CancelButtonLabel: 'Annuler',
    ThankYouTitle: 'Merci pour votre avis !',
    ThankYouSubtitle: 'Votre contribution nous aide à améliorer continuellement 3M Go.'
  },
  it: {
    YesButtonLabel: 'Sì',
    NoButtonLabel: 'No',
    CommentPlaceholder: '(Facoltativo) Fornisci un commento per aiutarci a migliorare continuamente 3M Go.',
    SubmitButtonLabel: 'Invia',
    SubmittingButtonLabel: 'Invio in corso…',
    CancelButtonLabel: 'Annulla',
    ThankYouTitle: 'Grazie per il tuo feedback!',
    ThankYouSubtitle: 'Il tuo contributo ci aiuta a migliorare continuamente 3M Go.'
  },
  nl: {
    YesButtonLabel: 'Ja',
    NoButtonLabel: 'Nee',
    CommentPlaceholder: '(Optioneel) Geef feedback zodat we 3M Go voortdurend kunnen verbeteren.',
    SubmitButtonLabel: 'Verzenden',
    SubmittingButtonLabel: 'Bezig met verzenden…',
    CancelButtonLabel: 'Annuleren',
    ThankYouTitle: 'Bedankt voor uw feedback!',
    ThankYouSubtitle: 'Uw inbreng helpt ons 3M Go voortdurend te verbeteren.'
  },
  ja: {
    YesButtonLabel: 'はい',
    NoButtonLabel: 'いいえ',
    CommentPlaceholder: '(任意) 3M Goを継続的に改善するためのご意見をお聞かせください。',
    SubmitButtonLabel: '送信',
    SubmittingButtonLabel: '送信中…',
    CancelButtonLabel: 'キャンセル',
    ThankYouTitle: 'フィードバックをありがとうございます！',
    ThankYouSubtitle: 'いただいたご意見は3M Goの継続的な改善に役立てられます。'
  },
  ko: {
    YesButtonLabel: '예',
    NoButtonLabel: '아니요',
    CommentPlaceholder: '(선택 사항) 3M Go를 지속적으로 개선할 수 있도록 의견을 남겨주세요.',
    SubmitButtonLabel: '제출',
    SubmittingButtonLabel: '제출 중…',
    CancelButtonLabel: '취소',
    ThankYouTitle: '피드백을 보내주셔서 감사합니다!',
    ThankYouSubtitle: '고객님의 의견은 3M Go를 지속적으로 개선하는 데 도움이 됩니다.'
  },
  pl: {
    YesButtonLabel: 'Tak',
    NoButtonLabel: 'Nie',
    CommentPlaceholder: '(Opcjonalnie) Prosimy o przekazanie opinii, abyśmy mogli stale ulepszać 3M Go.',
    SubmitButtonLabel: 'Wyślij',
    SubmittingButtonLabel: 'Wysyłanie…',
    CancelButtonLabel: 'Anuluj',
    ThankYouTitle: 'Dziękujemy za opinię!',
    ThankYouSubtitle: 'Twoja opinia pomaga nam stale ulepszać 3M Go.'
  },
  pt: {
    YesButtonLabel: 'Sim',
    NoButtonLabel: 'Não',
    CommentPlaceholder: '(Opcional) Forneça seu feedback para que possamos melhorar continuamente o 3M Go.',
    SubmitButtonLabel: 'Enviar',
    SubmittingButtonLabel: 'Enviando…',
    CancelButtonLabel: 'Cancelar',
    ThankYouTitle: 'Obrigado pelo seu feedback!',
    ThankYouSubtitle: 'Sua opinião nos ajuda a melhorar continuamente o 3M Go.'
  },
  th: {
    YesButtonLabel: 'ใช่',
    NoButtonLabel: 'ไม่ใช่',
    CommentPlaceholder: '(ไม่บังคับ) กรุณาแสดงความคิดเห็นเพื่อให้เราปรับปรุง 3M Go อย่างต่อเนื่อง',
    SubmitButtonLabel: 'ส่ง',
    SubmittingButtonLabel: 'กำลังส่ง…',
    CancelButtonLabel: 'ยกเลิก',
    ThankYouTitle: 'ขอบคุณสำหรับความคิดเห็นของคุณ!',
    ThankYouSubtitle: 'ความคิดเห็นของคุณช่วยให้เราปรับปรุง 3M Go อย่างต่อเนื่อง'
  },
  'zh-cn': {
    YesButtonLabel: '是',
    NoButtonLabel: '否',
    CommentPlaceholder: '（可选）请提供您的反馈，以便我们持续改进3M Go。',
    SubmitButtonLabel: '提交',
    SubmittingButtonLabel: '提交中…',
    CancelButtonLabel: '取消',
    ThankYouTitle: '感谢您的反馈！',
    ThankYouSubtitle: '您的意见有助于我们持续改进3M Go。'
  },
  'zh-tw': {
    YesButtonLabel: '是',
    NoButtonLabel: '否',
    CommentPlaceholder: '（選填）請提供您的意見，以便我們持續改進3M Go。',
    SubmitButtonLabel: '提交',
    SubmittingButtonLabel: '提交中…',
    CancelButtonLabel: '取消',
    ThankYouTitle: '感謝您的意見反饋！',
    ThankYouSubtitle: '您的意見有助於我們持續改進3M Go。'
  }
};

// Simple lookup, English is default if someone passes wrong/unknown key
export function getFeedbackStrings(language: string): IFeedbackStrings {
  const languageKey = (language || 'en').toLowerCase();
  return STRINGS[languageKey] || STRINGS.en;
}
