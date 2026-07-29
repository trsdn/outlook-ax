import Foundation

// MARK: - L10n Catalog
//
// One shared catalog for all localized AX label matching.
// Arrays follow the convention: [de, en, fr, es, it, ...]
// Command-local localized strings are prohibited — add them here instead.
//
// Adding a new locale: append its label to each relevant array.
// Fixture-verified variants required for: de, en, fr, es, it

public enum L10n {

    // MARK: - Window titles

    public static let calendarWindow       = ["Kalender", "Calendar", "Calendrier", "Calendario"]
    public static let inboxWindow          = ["Posteingang", "Inbox", "Boîte de réception", "Bandeja de entrada", "Posta in arrivo"]

    // MARK: - Calendar: table / list view

    public static let calendarEventsTable  = ["Kalenderereignisse", "Calendar events", "Événements du calendrier", "Eventos del calendario", "Eventi del calendario"]
    public static let allDay               = ["ganztägig", "all day", "all-day", "toute la journée", "todo el día", "tutto il giorno"]

    // MARK: - Calendar: status (cellDesc suffix after showAsPrefix)

    public static let showAsPrefix         = ["anzeigen als", "show as", "afficher comme", "mostrar como", "mostra come"]
    public static let statusBusy           = ["Gebucht", "Busy", "Occupé", "Ocupado", "Occupato"]
    public static let statusFree           = ["Frei", "Free", "Disponible", "Libre", "Libero"]
    public static let statusTentative      = ["Mit Vorbehalt", "Tentative", "Provisoire", "Provisional", "Provvisorio"]
    public static let statusOOF            = ["Außer Haus", "Out of Office", "Absent(e)", "Fuera de la oficina", "Fuori ufficio"]
    public static let statusElsewhere      = ["An anderem Ort", "Working Elsewhere", "Travaille ailleurs", "Trabajando en otro lugar", "Lavora altrove"]

    // MARK: - Calendar: organizer

    public static let organizerPrefix      = ["Organisator ", "Organizer ", "Organisateur ", "Organizador ", "Organizzatore "]
    public static let youAreOrganizer      = ["Sie sind der Organisator", "You are the organizer", "Vous êtes l'organisateur", "Usted es el organizador", "Sei l'organizzatore"]

    // MARK: - Calendar: categories

    public static let category             = ["Kategorie", "Category", "Catégorie", "Categoría", "Categoria"]

    // MARK: - Calendar: detail window sections

    public static let sectionOrganizer     = ["Organisator", "Organizer", "Organisateur", "Organizador", "Organizzatore"]
    public static let sectionRequired      = ["Erforderlich", "Required", "Obligatoire", "Obligatorio", "Obbligatorio"]
    public static let sectionOptional      = ["Optional", "Facultatif", "Opcional", "Facoltativo"]

    // MARK: - Calendar: attendee responses (suffixes)

    public static let respAccepted         = ["angenommen.", "accepted."]
    public static let respDeclined         = ["abgesagt.", "declined."]
    public static let respNotResponded     = ["nicht geantwortet.", "not responded.", "haven't responded."]
    public static let respTentative        = ["mit Vorbehalt.", "tentative.", "tentatively."]

    // MARK: - Calendar: detail window noise filters

    public static let detailNoise          = ["statt.", "Findet am", "Takes place", "instead.", "Tiene lugar"]

    // MARK: - Calendar: title prefixes for myResponse

    public static let declinedPrefix       = ["Declined: ", "Abgelehnt: "]
    public static let followingPrefix      = ["Following: "]

    // MARK: - Calendar: navigation buttons

    public static let today                = ["Heute", "Today", "Aujourd'hui", "Hoy", "Oggi"]
    public static let nextDay              = ["Nächster Tag", "Next day", "Next Day", "Jour suivant", "Día siguiente", "Giorno successivo"]
    public static let prevDay              = ["Vorheriger Tag", "Previous day", "Previous Day", "Jour précédent", "Día anterior", "Giorno precedente"]

    // MARK: - Calendar: create event

    public static let newEvent             = ["Neuer Termin", "New event", "New Event", "Nouvel événement", "Nuevo evento", "Nuovo evento"]
    public static let subject              = ["Betreff", "Subject", "Objet", "Asunto", "Oggetto"]
    public static let addAttendees         = ["Erforderliche Personen hinzufügen", "Erforderliche Teilnehmer",
                                              "Add required attendees", "Required attendees",
                                              "Ajouter des participants obligatoires",
                                              "Agregar asistentes requeridos",
                                              "Aggiungere i partecipanti richiesti"]
    public static let startDate            = ["Beginnt am", "Starts on", "Start date", "Commence le", "Fecha de inicio", "Data di inizio"]
    public static let startTime            = ["Beginnt um", "Starts at", "Start time", "Commence à", "Hora de inicio", "Ora di inizio"]
    public static let send                 = ["Senden", "Send", "Envoyer", "Enviar", "Invia"]
    public static let save                 = ["Speichern", "Save", "Enregistrer", "Guardar", "Salva"]
    public static let discard              = ["Verwerfen", "Discard", "Ignorer", "Descartar", "Elimina"]

    // MARK: - Calendar: view picker

    public static let calendarViewPicker   = ["Kalenderansicht", "Calendar view", "Vue du calendrier", "Vista del calendario"]

    // MARK: - Calendar: sidebar / nav pane

    public static let navPane              = ["Navigationsbereich", "Navigation pane", "Volet de navigation"]
    public static let myCalendars          = ["Meine Kalender", "My Calendars", "Mes calendriers", "Mis calendarios", "Calendari personali"]
    public static let otherCalendars       = ["Andere Kalender", "Other Calendars", "Autres calendriers", "Otros calendarios", "Altri calendari"]
    public static let calendarShown        = ["Angezeigt", "Shown", "Displayed", "Teilweise angezeigt", "Partially shown"]

    // MARK: - Calendar: event editor window title fragments

    public static let eventWindowTitles    = ["Termin", "Event", "Événement", "Evento", "Appuntamento"]

    // MARK: - Mail: header

    public static let messageHeader        = ["Nachrichtenkopfzeile", "Message header", "En-tête du message"]
    public static let headerDetails        = ["Nachrichtenkopfdetails", "Message header details"]
    public static let sentPrefix           = ["Gesendet am:", "Gesendet", "Sent on:", "Sent"]
    public static let messageList          = ["Nachrichtenliste", "Message list", "Liste de messages", "Lista de mensajes", "Elenco messaggi"]

    // MARK: - Mail: compose

    public static let newEmail             = ["Neue E-Mail", "New Email", "New email", "Nouveau courrier", "Nuevo correo electrónico", "Nuovo messaggio"]
    public static let composeWindow        = ["Nachricht", "Message", "Neue", "New"]
    public static let toField              = ["An", "To", "À", "Empfänger", "Recipients", "Para"]
    public static let bodyField            = ["Nachrichtentext", "Message body"]
    public static let fromPrefix           = ["Von: ", "From: ", "De : ", "De: ", "Da: "]

    // MARK: - Mail: search

    public static let search               = ["Suchen", "Search", "Rechercher", "Buscar", "Cerca"]

    // MARK: - Mail: actions

    public static let reply                = ["Antworten", "Reply", "Répondre", "Responder", "Rispondi"]
    public static let replyAll             = ["Allen antworten", "Reply All", "Répondre à tous", "Responder a todos", "Rispondi a tutti"]
    public static let forward              = ["Weiterleiten", "Forward", "Transférer", "Reenviar", "Inoltra"]
    public static let delete               = ["Löschen", "Delete", "Supprimer", "Eliminar", "Elimina"]
    public static let archive              = ["Archivieren", "Archive", "Archiver", "Archivar", "Archivia"]
    public static let flag                 = ["Kennzeichnen", "Flag", "Marquer", "Marcar", "Contrassegna"]
    public static let markRead             = ["Markieren:", "Mark:"]
    public static let move                 = ["Verschieben", "Move", "Déplacer", "Mover", "Sposta"]
    public static let report               = ["Melden", "Report", "Signaler", "Informar", "Segnala"]
    public static let react                = ["Reagieren", "React", "Réagir", "Reaccionar", "Reazione"]
    public static let summarize            = ["Zusammenfassen", "Summarize", "Résumer", "Resumir", "Riassumi"]
    public static let filterSort           = ["Filtern und sortieren", "Filter and sort", "Filtrer et trier"]
    public static let moreItems            = ["Weitere Elemente anzeigen", "Show more items"]

    // MARK: - Navigation

    public static let navCalendar          = ["Kalender", "Calendar", "Calendrier", "Calendario"]
    public static let navMail              = ["E-Mail", "Mail", "Courrier", "Correo"]
    public static let navPeople            = ["Personen", "People", "Contacts", "Contactos", "Contatti"]
    public static let navTasks             = ["Aufgaben", "Tasks", "Tâches", "Tareas", "Attività"]
    public static let navCopilot           = ["Copilot"]
    public static let navOneDrive          = ["OneDrive"]
    public static let navFavorites         = ["Favoriten", "Favorites", "Favoris", "Favoritos"]
    public static let navOrgExplorer       = ["Organisations-Explorer", "Org Explorer"]

    // MARK: - Notifications

    public static let newNotifications     = ["Neue Benachrichtigungen:", "New notifications:", "Nouvelles notifications:", "Nuevas notificaciones:"]

    // MARK: - My Day

    public static let myDay                = ["Mein Tag", "My Day", "Ma journée", "Mi día", "La mia giornata"]

    // MARK: - Menu paths

    public static let menuView             = ["Anzeigen", "View", "Affichage", "Ver"]
    public static let menuSwitchTo         = ["Wechseln zu", "Switch to", "Basculer vers", "Cambiar a"]
    public static let menuEvent            = ["Ereignis", "Event", "Événement", "Evento"]
    public static let menuTools            = ["Werkzeuge", "Tools", "Outils", "Herramientas"]
    public static let menuSync             = ["Synchronisieren", "Sync", "Synchroniser", "Sincronizar"]
    public static let menuAutoReply        = ["Automatische Antworten...", "Automatic Replies...", "Réponses automatiques...", "Respuestas automáticas..."]
    public static let menuShowAs           = ["Anzeigen als", "Show As"]
    public static let menuCategorize       = ["Kategorisieren", "Categorize", "Catégoriser", "Categorizar"]
    public static let menuPrivate          = ["Privat", "Private", "Privé", "Privado", "Privato"]

    // MARK: - Calendar: view modes

    public static let viewDay              = ["Tag", "Day", "Jour", "Día", "Giorno"]
    public static let viewWorkWeek         = ["Arbeitswoche", "Work Week", "Semaine de travail", "Semana laboral", "Settimana lavorativa"]
    public static let viewWeek             = ["Woche", "Week", "Semaine", "Semana", "Settimana"]
    public static let viewMonth            = ["Monat", "Month", "Mois", "Mes", "Mese"]
    public static let viewThreeDay         = ["Drei Tage", "Three Day", "Trois jours", "Tres días", "Tre giorni"]
    public static let viewList             = ["Liste", "List", "Elenco"]

    // MARK: - Calendar: timescale

    public static let minutesSuffix        = ["Minuten", "Minutes", "minutes"]

    // MARK: - Calendar: filter

    public static let filterAll            = ["Alle", "All", "Tous", "Todos", "Tutti"]
    public static let filterAppointments   = ["Termine", "Appointments", "Rendez-vous", "Citas", "Appuntamenti"]
    public static let filterMeetings       = ["Besprechungen", "Meetings", "Réunions", "Reuniones", "Riunioni"]
    public static let filterCategories     = ["Kategorien", "Categories", "Catégories", "Categorías", "Categorie"]
    public static let filterShowAs         = ["Anzeigen als", "Show As"]
    public static let filterRecurring      = ["Wiederholung", "Recurring", "Récurrence", "Recurrentes", "Ricorrenti"]
    public static let filterPrivacy        = ["Datenschutz", "Privacy", "Confidentialité", "Privacidad", "Privacy"]
    public static let filterDeclined       = ["Abgelehnte Ereignisse ausblenden", "Hide declined events"]

    // MARK: - Calendar: event menu actions

    public static let accept               = ["Akzeptieren", "Zusagen", "Accept", "Accepter", "Aceptar", "Accetta"]
    public static let tentative            = ["Mit Vorbehalt", "Tentative", "Provisoire", "Provisional", "Provvisorio"]
    public static let decline              = ["Ablehnen", "Decline", "Refuser", "Rechazar", "Rifiuta"]
    public static let joinMeeting          = ["An Onlinebesprechung teilnehmen", "Join Online Meeting", "Rejoindre la réunion", "Unirse a la reunión en línea"]
    public static let duplicateEvent       = ["Ereignis duplizieren", "Duplicate Event", "Dupliquer l'événement", "Duplicar evento"]
    public static let cancelMeeting        = ["Besprechung absagen", "Cancel Meeting", "Annuler la réunion", "Cancelar reunión"]

    // MARK: - Calendar: show-as menu values

    public static let showAsFree           = ["Frei", "Free", "Disponible", "Libre", "Libero"]
    public static let showAsTentative      = ["Mit Vorbehalt", "Tentative", "Provisoire", "Provisional", "Provvisorio"]
    public static let showAsBusy           = ["Gebucht", "Busy", "Occupé", "Ocupado", "Occupato"]
    public static let showAsOOF            = ["Außer Haus", "Out of Office", "Absent(e)", "Fuera de la oficina", "Fuori ufficio"]
    public static let showAsElsewhere      = ["An anderem Ort tätig", "Working Elsewhere", "Travaille ailleurs", "Trabajando en otro lugar"]

    // MARK: - Calendar: color menu

    public static let colorBlue            = ["Blau", "Blue", "Bleu", "Azul", "Blu"]
    public static let colorGreen           = ["Grün", "Green", "Vert", "Verde"]
    public static let colorOrange          = ["Orange", "Naranja", "Arancione"]
    public static let colorPlatinum        = ["Platingrau", "Platinum", "Platine", "Platino"]
    public static let colorYellow          = ["Gelb", "Yellow", "Jaune", "Amarillo", "Giallo"]
    public static let colorCyan            = ["Zyan", "Cyan"]
    public static let colorMagenta         = ["Magenta"]
    public static let colorBrown           = ["Braun", "Brown", "Marron", "Marrón", "Marrone"]
    public static let colorBurgundy        = ["Burgunderrot", "Burgundy", "Bordeaux", "Burdeos"]
    public static let colorTeal            = ["Meeresgrün", "Teal", "Sarcelle"]
    public static let colorLilac           = ["Flieder", "Lilac", "Lilas", "Lila"]

    // MARK: - Calendar: timescale menu

    public static let menuTimescale        = ["Zeitskala", "Timescale", "Échelle de temps", "Escala de tiempo"]
    public static let menuFilter           = ["Filtern", "Filter", "Filtrer", "Filtrar"]
    public static let menuColor            = ["Farbe", "Color", "Couleur"]

    // MARK: - Mail: folder sidebar

    public static let favorites            = ["Favoriten", "Favorites", "Favoris"]
    public static let allAccounts          = ["Alle Konten", "All Accounts", "Tous les comptes"]
    public static let groups               = ["Gruppen", "Groups", "Groupes"]

    // MARK: - L10n string matching helpers

    /// True when `text` contains any of the variants.
    public static func matches(_ text: String, _ variants: [String]) -> Bool {
        variants.contains(where: { text.contains($0) })
    }

    /// True when `text` starts with any of the variants.
    public static func startsWith(_ text: String, _ variants: [String]) -> Bool {
        variants.contains(where: { text.hasPrefix($0) })
    }

    /// True when `text` exactly equals any of the variants.
    public static func equals(_ text: String, _ variants: [String]) -> Bool {
        variants.contains(text)
    }

    /// True when `text` ends with any of the variants.
    public static func endsWith(_ text: String, _ variants: [String]) -> Bool {
        variants.contains(where: { text.hasSuffix($0) })
    }
}
