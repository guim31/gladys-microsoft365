// -----------------------------------------------------------------------------
// Every text the code produces, in the four languages of the Gladys interface.
// German says "du", Spanish "usted" (test/i18n.test.js checks the German).
// Brands (Microsoft, Outlook, To Do) are never translated.
// -----------------------------------------------------------------------------

import { LANGUAGES } from './config.js';

export const LOCALES = { en: 'en-GB', fr: 'fr-FR', de: 'de-DE', es: 'es-ES' };

/** Feature names, frozen when the user adds the device (config `language`). */
export const NAMES = {
  en: {
    busy: 'Busy',
    out_of_office: 'Out of office',
    current_event: 'Current event',
    next_event: 'Next event',
    open_tasks: 'Open tasks',
    overdue_tasks: 'Overdue tasks',
    due_today_tasks: 'Tasks due today',
  },
  fr: {
    busy: 'Occupé',
    out_of_office: 'Absent du bureau',
    current_event: 'Évènement en cours',
    next_event: 'Prochain évènement',
    open_tasks: 'Tâches ouvertes',
    overdue_tasks: 'Tâches en retard',
    due_today_tasks: "Tâches pour aujourd'hui",
  },
  de: {
    busy: 'Beschäftigt',
    out_of_office: 'Abwesend',
    current_event: 'Aktueller Termin',
    next_event: 'Nächster Termin',
    open_tasks: 'Offene Aufgaben',
    overdue_tasks: 'Überfällige Aufgaben',
    due_today_tasks: 'Heute fällige Aufgaben',
  },
  es: {
    busy: 'Ocupado',
    out_of_office: 'Fuera de la oficina',
    current_event: 'Evento en curso',
    next_event: 'Próximo evento',
    open_tasks: 'Tareas abiertas',
    overdue_tasks: 'Tareas vencidas',
    due_today_tasks: 'Tareas para hoy',
  },
};

/** Short texts: feature values, widget labels and buttons. */
export const TEXT = {
  en: {
    none: 'None',
    today: 'today',
    tomorrow: 'tomorrow',
    allDay: 'all day',
    now: 'Now',
    todayAgenda: 'Today',
    noEventToday: 'Nothing else in the calendar today.',
    noOpenTask: 'No open task.',
    noDueDate: 'no due date',
    overdue: 'overdue',
    open: 'Open',
    late: 'Overdue',
    addTask: 'Add a task',
    taskTitle: 'Task',
    due: 'Due',
    dueNone: 'No due date',
    dueToday: 'Today',
    dueTomorrow: 'Tomorrow',
    dueWeek: 'In a week',
    notSignedIn: 'Not signed in to Microsoft: use “Sign in” in the integration settings.',
    nothingToShow: 'Add an Outlook calendar or a To Do list to Gladys first (“Discover” tab).',
    syncing: 'First read of Microsoft 365 in progress…',
  },
  fr: {
    none: 'Aucun',
    today: "aujourd'hui",
    tomorrow: 'demain',
    allDay: 'toute la journée',
    now: 'En cours',
    todayAgenda: "Aujourd'hui",
    noEventToday: "Plus rien à l'agenda aujourd'hui.",
    noOpenTask: 'Aucune tâche ouverte.',
    noDueDate: 'sans échéance',
    overdue: 'en retard',
    open: 'Ouvertes',
    late: 'En retard',
    addTask: 'Ajouter une tâche',
    taskTitle: 'Tâche',
    due: 'Échéance',
    dueNone: 'Sans échéance',
    dueToday: "Aujourd'hui",
    dueTomorrow: 'Demain',
    dueWeek: 'Dans une semaine',
    notSignedIn:
      "Pas connecté à Microsoft : utilisez « Se connecter » dans les réglages de l'intégration.",
    nothingToShow:
      "Ajoutez d'abord un calendrier Outlook ou une liste To Do à Gladys (onglet « Découverte »).",
    syncing: 'Première lecture de Microsoft 365 en cours…',
  },
  de: {
    none: 'Keiner',
    today: 'heute',
    tomorrow: 'morgen',
    allDay: 'ganztägig',
    now: 'Jetzt',
    todayAgenda: 'Heute',
    noEventToday: 'Heute steht nichts mehr im Kalender.',
    noOpenTask: 'Keine offene Aufgabe.',
    noDueDate: 'ohne Fälligkeit',
    overdue: 'überfällig',
    open: 'Offen',
    late: 'Überfällig',
    addTask: 'Aufgabe hinzufügen',
    taskTitle: 'Aufgabe',
    due: 'Fällig',
    dueNone: 'Ohne Fälligkeit',
    dueToday: 'Heute',
    dueTomorrow: 'Morgen',
    dueWeek: 'In einer Woche',
    notSignedIn:
      'Nicht bei Microsoft angemeldet: Nutze „Anmelden“ in den Einstellungen der Integration.',
    nothingToShow:
      'Füge zuerst einen Outlook-Kalender oder eine To-Do-Liste zu Gladys hinzu (Tab „Entdecken“).',
    syncing: 'Microsoft 365 wird zum ersten Mal gelesen…',
  },
  es: {
    none: 'Ninguno',
    today: 'hoy',
    tomorrow: 'mañana',
    allDay: 'todo el día',
    now: 'En curso',
    todayAgenda: 'Hoy',
    noEventToday: 'Nada más en el calendario hoy.',
    noOpenTask: 'Ninguna tarea abierta.',
    noDueDate: 'sin vencimiento',
    overdue: 'vencida',
    open: 'Abiertas',
    late: 'Vencidas',
    addTask: 'Añadir una tarea',
    taskTitle: 'Tarea',
    due: 'Vencimiento',
    dueNone: 'Sin vencimiento',
    dueToday: 'Hoy',
    dueTomorrow: 'Mañana',
    dueWeek: 'Dentro de una semana',
    notSignedIn:
      'Sin conexión con Microsoft: use «Iniciar sesión» en los ajustes de la integración.',
    nothingToShow:
      'Añada primero un calendario de Outlook o una lista de To Do a Gladys (pestaña «Descubrir»).',
    syncing: 'Primera lectura de Microsoft 365 en curso…',
  },
};

/**
 * Messages: action results, connection status, scene and widget errors.
 * Functions take the details to insert.
 */
export const MESSAGES = {
  en: {
    missingClientId:
      'Enter the application (client) ID of your Azure app registration first, then save.',
    signInCode: (uri, code, minutes) =>
      `Open ${uri} and enter the code ${code} within ${minutes} minutes, then sign in with your Microsoft account.`,
    signInPending: (uri, code) => `Waiting for sign-in: open ${uri} and enter the code ${code}.`,
    signedIn: (name) => `Signed in as ${name}.`,
    signInExpired: 'The sign-in code expired: run “Sign in” again.',
    signInDeclined: 'Sign-in declined on the Microsoft page.',
    signInRefused: (code) =>
      `Microsoft refused the sign-in (${code}): check the client ID, the tenant and “Allow public client flows”.`,
    notSignedIn: 'Not signed in: use “Sign in”.',
    sessionExpired: 'The Microsoft session expired or was revoked: use “Sign in” again.',
    signedOut: 'Signed out: the Microsoft tokens were deleted.',
    unreachable: 'Microsoft 365 is unreachable for now, the integration retries on its own.',
    throttled: (minutes) => `Microsoft 365 asks to slow down: next read in ${minutes} min.`,
    report: (name, calendars, lists) =>
      `Signed in as ${name}: ${calendars} calendar(s), ${lists} To Do list(s). Add them in the “Discover” tab.`,
    failed: (reason) => `Microsoft 365 answered with an error (${reason}).`,
    unknownList: 'Choose a To Do list added to Gladys.',
    unknownCalendar: 'Choose an Outlook calendar added to Gladys.',
    emptyTitle: 'The title is empty.',
    taskAdded: (title) => `Task “${title}” added.`,
  },
  fr: {
    missingClientId:
      "Saisissez d'abord l'ID d'application (client) de votre inscription d'application Azure, puis enregistrez.",
    signInCode: (uri, code, minutes) =>
      `Ouvrez ${uri} et saisissez le code ${code} dans les ${minutes} minutes, puis connectez-vous avec votre compte Microsoft.`,
    signInPending: (uri, code) =>
      `Connexion en attente : ouvrez ${uri} et saisissez le code ${code}.`,
    signedIn: (name) => `Connecté en tant que ${name}.`,
    signInExpired: 'Le code de connexion a expiré : relancez « Se connecter ».',
    signInDeclined: 'Connexion refusée sur la page Microsoft.',
    signInRefused: (code) =>
      `Microsoft a refusé la connexion (${code}) : vérifiez l'ID client, le locataire et « Autoriser les flux de clients publics ».`,
    notSignedIn: 'Pas connecté : utilisez « Se connecter ».',
    sessionExpired: 'La session Microsoft a expiré ou a été révoquée : relancez « Se connecter ».',
    signedOut: 'Déconnecté : les jetons Microsoft ont été supprimés.',
    unreachable: "Microsoft 365 est injoignable pour l'instant, l'intégration réessaie seule.",
    throttled: (minutes) =>
      `Microsoft 365 demande de ralentir : prochaine lecture dans ${minutes} min.`,
    report: (name, calendars, lists) =>
      `Connecté en tant que ${name} : ${calendars} calendrier(s), ${lists} liste(s) To Do. Ajoutez-les dans l'onglet « Découverte ».`,
    failed: (reason) => `Microsoft 365 a répondu par une erreur (${reason}).`,
    unknownList: 'Choisissez une liste To Do ajoutée à Gladys.',
    unknownCalendar: 'Choisissez un calendrier Outlook ajouté à Gladys.',
    emptyTitle: 'Le titre est vide.',
    taskAdded: (title) => `Tâche « ${title} » ajoutée.`,
  },
  de: {
    missingClientId:
      'Gib zuerst die Anwendungs-ID (Client) deiner Azure-App-Registrierung ein und speichere.',
    signInCode: (uri, code, minutes) =>
      `Öffne ${uri}, gib innerhalb von ${minutes} Minuten den Code ${code} ein und melde dich mit deinem Microsoft-Konto an.`,
    signInPending: (uri, code) =>
      `Anmeldung ausstehend: Öffne ${uri} und gib den Code ${code} ein.`,
    signedIn: (name) => `Angemeldet als ${name}.`,
    signInExpired: 'Der Anmeldecode ist abgelaufen: Starte „Anmelden“ erneut.',
    signInDeclined: 'Anmeldung auf der Microsoft-Seite abgelehnt.',
    signInRefused: (code) =>
      `Microsoft hat die Anmeldung abgelehnt (${code}): Prüfe die Client-ID, den Mandanten und „Öffentliche Clientflows zulassen“.`,
    notSignedIn: 'Nicht angemeldet: Nutze „Anmelden“.',
    sessionExpired:
      'Die Microsoft-Sitzung ist abgelaufen oder wurde widerrufen: Nutze „Anmelden“ erneut.',
    signedOut: 'Abgemeldet: Die Microsoft-Token wurden gelöscht.',
    unreachable:
      'Microsoft 365 ist gerade nicht erreichbar, die Integration versucht es selbst erneut.',
    throttled: (minutes) =>
      `Microsoft 365 bittet um weniger Anfragen: nächster Abruf in ${minutes} Min.`,
    report: (name, calendars, lists) =>
      `Angemeldet als ${name}: ${calendars} Kalender, ${lists} To-Do-Liste(n). Füge sie im Tab „Entdecken“ hinzu.`,
    failed: (reason) => `Microsoft 365 hat mit einem Fehler geantwortet (${reason}).`,
    unknownList: 'Wähle eine zu Gladys hinzugefügte To-Do-Liste.',
    unknownCalendar: 'Wähle einen zu Gladys hinzugefügten Outlook-Kalender.',
    emptyTitle: 'Der Titel ist leer.',
    taskAdded: (title) => `Aufgabe „${title}“ hinzugefügt.`,
  },
  es: {
    missingClientId:
      'Introduzca primero el ID de aplicación (cliente) de su registro de aplicación de Azure y guarde.',
    signInCode: (uri, code, minutes) =>
      `Abra ${uri} e introduzca el código ${code} antes de ${minutes} minutos; luego inicie sesión con su cuenta de Microsoft.`,
    signInPending: (uri, code) =>
      `Inicio de sesión pendiente: abra ${uri} e introduzca el código ${code}.`,
    signedIn: (name) => `Sesión iniciada como ${name}.`,
    signInExpired: 'El código de inicio de sesión ha caducado: vuelva a usar «Iniciar sesión».',
    signInDeclined: 'Inicio de sesión rechazado en la página de Microsoft.',
    signInRefused: (code) =>
      `Microsoft ha rechazado el inicio de sesión (${code}): compruebe el ID de cliente, el inquilino y «Permitir flujos de clientes públicos».`,
    notSignedIn: 'Sin sesión iniciada: use «Iniciar sesión».',
    sessionExpired:
      'La sesión de Microsoft ha caducado o ha sido revocada: vuelva a usar «Iniciar sesión».',
    signedOut: 'Sesión cerrada: los tokens de Microsoft se han eliminado.',
    unreachable:
      'Microsoft 365 no está accesible por ahora, la integración lo reintentará por sí sola.',
    throttled: (minutes) =>
      `Microsoft 365 pide reducir las peticiones: próxima lectura dentro de ${minutes} min.`,
    report: (name, calendars, lists) =>
      `Sesión iniciada como ${name}: ${calendars} calendario(s), ${lists} lista(s) de To Do. Añádalos en la pestaña «Descubrir».`,
    failed: (reason) => `Microsoft 365 ha respondido con un error (${reason}).`,
    unknownList: 'Elija una lista de To Do añadida a Gladys.',
    unknownCalendar: 'Elija un calendario de Outlook añadido a Gladys.',
    emptyTitle: 'El título está vacío.',
    taskAdded: (title) => `Tarea «${title}» añadida.`,
  },
};

/** The table of one language, English when unknown. */
export function pick(table, language) {
  return table[language] ?? table.en;
}

/**
 * A multi-language object from a message of MESSAGES, e.g.
 * `everyLanguage('signedIn', 'Jane')` → `{ en: 'Signed in as Jane.', fr: … }`.
 */
export function everyLanguage(key, ...args) {
  return Object.fromEntries(
    LANGUAGES.map((language) => {
      const message = MESSAGES[language][key];
      return [language, typeof message === 'function' ? message(...args) : message];
    }),
  );
}

/** A short text of TEXT in every language (form labels of a widget button). */
export function textInEveryLanguage(key) {
  return Object.fromEntries(LANGUAGES.map((language) => [language, TEXT[language][key]]));
}

/** An Error whose message is in the configured language. */
export function localizedError(language, key, ...args) {
  const message = pick(MESSAGES, language)[key];
  return new Error(typeof message === 'function' ? message(...args) : message);
}
