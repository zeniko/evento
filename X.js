﻿/* Evento Excel-Helfer  (C) 2009 - 2019 Simon Bünzli  <simon.buenzli@zeniko.ch> */

/*
Zum Testen die folgende Zeile in die Adresszeile des Browsers kopieren:

javascript:void(document.body.appendChild(document.createElement("script")).src="https://www.zeniko.ch/evento/X.js")
*/

// jQuery nachladen, sofern noch nicht verfügbar (wird benötigt)
if (!window.jQuery) {
    document.body.appendChild(document.createElement("script")).src = "https://www.zeniko.ch/evento/jquery.min.js";
}

// beim Neuladen des Helfers den bereits geladenen Helfer zuerst deaktivieren
if (window.X && window.X.uninit) {
    X.uninit();
}

// Namespace für sämtliche zusätzliche Funktionalität
var X = {
    // Version des Scripts:
    version: "0.6", // Stand 05.08.19

    // Wenn X.js eingebettet ist, erscheint das Overlay am Anfang nicht
    // das Callback wird am Ende von onFrameLoad aufgerufen
    // (siehe EVT CR-11450)
    embedded: window.X_IS_EMBEDDED || false,
    embeddedCallback: window.X_EMBEDDED_CALLBACK || function() {},

    // im Interface verwendete Sprache (muss in X.strings vorhanden sein)
    lang: "de",

    // im Interface sichtbare Texte
    // Text von start_button wurde angepasst für EVT CR-11450
    strings: {
        de: {
            views: [{
                start_dropdown: "Excel-Eingabe",
                start_button: "für Noten",
                accept_button: "Noten übernehmen",
                cancel_button: "Abbrechen",
                feedback_to: "Feedback an %s", // %s wird durch eine E-Mail-Adresse ersetzt
                default_lines: [
                    "# Hierhin können Daten aus einer Tabelle kopiert/eingefügt werden",
                    "# (für Kurs %s):", // %s wird durch die Kursbezeichnung ersetzt
                    "",
                    "# die ERSTE Spalte (oder die ersten zwei Spalten) müssen",
                    "# die Namen der SchülerInnen enthalten, die LETZTE Spalte",
                    "# die Zeugnisnoten, dazwischen liegende Spalten werden ignoriert",
                    ""
                ]
            }, {
                start_dropdown: "Excel-Eingabe",
                start_button: "für Absenzen",
                accept_button: "Absenzen übernehmen",
                cancel_button: "Abbrechen",
                feedback_to: "Feedback an %s", // %s wird durch eine E-Mail-Adresse ersetzt
                default_lines: [
                    "# Hierhin können Daten aus einer Tabelle kopiert/eingefügt werden",
                    "",
                    "# die ERSTE Spalte (oder die ersten zwei Spalten) müssen",
                    "# die Namen der SchülerInnen enthalten, die ZWEI LETZTEN Spalten",
                    "# die Absenzen (zuerst die entschuldigten, dann die unentschuldigten)",
                    ""
                ]
            }],
            // Fehlermeldungen erscheinen direkt neben der Noten-/Absenzeneingabe:
            errors: {
                not_found: "Namen nicht gefunden",
                grade_not_found: "Note nicht gefunden",
                name_double: "Name erscheint mehrfach",
                invalid_value: "Ungültiger Wert: %s", // %s wird durch die ungültige Eingabe ersetzt
                no_number: "Keine Zahl?"
            }
        },
        fr: {
            views: [{
                start_dropdown: "Saisie Excel",
                start_button: "pour les notes",
                accept_button: "Valider les notes",
                cancel_button: "Annuler",
                feedback_to: "Envoyer un feedback à %s", // %s wird durch eine E-Mail-Adresse ersetzt
                default_lines: [
                    "# Possibilité de copier/insérer ici les données d’un tableau",
                    "# (pour le cours %s) :", // %s wird durch die Kursbezeichnung ersetzt
                    "",
                    "# La PREMIERE colonne ou les deux premières colonnes doivent",
                    "# contenir le nom des élèves, la DERNIERE colonne les notes",
                    "# de bulletin ; ignorer les colonnes situées entre celles-ci.",
                    ""
                ]
            }, {
                start_dropdown: "Saisie Excel",
                start_button: "pour les absences",
                accept_button: "Valider les absences",
                cancel_button: "Annuler",
                feedback_to: "Envoyer un feedback à %s", // %s wird durch eine E-Mail-Adresse ersetzt
                default_lines: [
                    "# Possibilité de copier/insérer ici les données d’un tableau :",
                    "",
                    "# La PREMIERE colonne ou les deux premières colonnes doivent",
                    "# contenir le nom des élèves, les DEUX DERNIERES colonnes",
                    "# les absences (d’abord les absences excusées, puis les autres)",
                    ""
                ]
            }],
            // Fehlermeldungen erscheinen direkt neben der Noten-/Absenzeneingabe:
            errors: {
                not_found: "Nom non trouvé",
                grade_not_found: "Note non trouvée",
                name_double: "Nom apparaissant plusieurs fois",
                invalid_value: "Valeur non valide : %s", // %s wird durch die ungültige Eingabe ersetzt
                no_number: "Aucun chiffre ?"
            }
        }
    },

    /**
     * initialisiert den Helfer
     */
    init: function() {
        if (X._loaded) {
            // versehentlich doppelt initialisiert?
            return;
        }

        if (!window.$) {
            // jQuery ist (noch) nicht bereit
            setTimeout(function() {
                X.init();
            }, 100);
            return;
        }

        X.parseNumber = X.memoizeFunction(X.parseNumber);
        X.unfancyName = X.memoizeFunction(X.unfancyName);
        X._loaded = true;

        var showPanel = !X.embedded;
        X.onFrameLoad(showPanel);
    },

    /**
     * räumt soweit auf, dass der Helfer ohne Performance-Einbusse neu geladen werden kann
     */
    uninit: function() {
        // entferne ggf. bereits eingefügte Elemente
        $("#tsv-overlay").remove();
    },

    /**
     * füge zusätzliche Eingabehilfen ein, sofern es solcher bedarf
     * @param aShowPanel  gibt an, ob das Excel-Importfeld unbedingt angezeigt werden soll
     *                    oder nur, wenn das Evento-Formular noch keine Daten enthält
     */
    onFrameLoad: function(aShowPanel) {
        if ($("td:contains('Anmeldungen'), th:contains('Anmeldungen')").length == 0) {
            X.lang = "fr";
        }
        X.strings[X.lang].views[2] = X.strings[X.lang].views[0];
        X.strings[X.lang].views[3] = X.strings[X.lang].views[1];

        var view = X.viewType();
        var strings = X.strings[X.lang].views[view];

        // füge den Knopf und das Textfeld (inkl. Styling) hinzu
        var pageEl = view == 2 ? $("div.page") : document.body;
        $(pageEl).append('\
<style type="text/css">\
	' + (view == 2 ? 'div.page { position: relative; }' : '') + ' \
	#tsv-overlay { position: fixed; top: 0px;' + (view != 2 ? 'left: 0px;' : 'max-width: 100%;') + 'width: 100%; height: 100%; display: none; } \
	#tsv-overlay-inner { height: 100%; background: white; padding: 5% 5% 20px; margin-top: 80px; margin-left: 300px; } \
	#tsv-overlay-inner-2 { height: 70%; } \
	/* Bugfix: Google Chrome ändert nur bei display:block Textfeldern mit CSS die Höhe */ \
	#tsv-data { width: 100%; height: 80%; margin-bottom: 1em; display: block; } \
	/* Bugfix: MSIE kennt "position: fixed" nicht */ \
	#tsv-overlay { _position: absolute; } \
</style>\
<div id="tsv-overlay"><div id="tsv-overlay-inner"><div id="tsv-overlay-inner-2">\
	<!-- Bugfix: MSIE7 kann im Standard Mode die Höhe von Textfeldern nicht mit CSS ändern -->\
	<textarea id="tsv-data" rows="20"></textarea>\
	\
	<div style="float: left;"><input type="button" value=" ' + strings.accept_button + ' " onclick="top.X.acceptOverlay(' + view + ');"> <input type="button" value=" ' + strings.cancel_button + ' " onclick="top.X.cancelOverlay();"></div>\
	<div style="float: right;">' + strings.feedback_to.replace("%s", '<a href="mailto:simon.buenzli@zeniko.ch?subject=Evento:%20Excel-Eingabe%20Feedback">Simon B&uuml;nzli</a>') + '</div>\
</div></div></div>\
		');

        if (view == 2) {
            // für JSModul sind Absenzen und Noten im gleichen Formular möglich
            if ($("td.gradeInput ~ td input[type=text]", pageEl).length >= 2) {
                var validGrades = X.collectValidGrades(view);
                // bei "besucht/dispensiert" Kursen werden manchmal bloss Absenzen eingegeben
                if (validGrades && validGrades.length == 2) {
                    view = 3;
                }
            }
            // für JSModul werden Links anstelle von Buttons verwendet
            $("#tsv-overlay input[type=button]", pageEl).replaceWith(function() {
                return '<a class="linkButton" onclick="' + this.getAttribute("onclick") + '" style="float: left; margin-right: 0.5em;"> ' + this.value + ' </a>';
            });
        }

        // lade das Excel-Importfeld automatisch, wenn noch keine Daten eingetragen sind
        var autoLoadOverlay = !X.embedded && (aShowPanel || $.grep(X.collectNames(view, true), function(aLine) {
            // enthält die Zeile bereits Daten (eine Note oder Absenzen)?
            return /\t/.test(aLine);
        }).length == 0);
        if (autoLoadOverlay) {
            X.showOverlay(view);
        }

        X.embeddedCallback();
    },

    /**
     * zeigt das Excel-Eingabefeld an (und füllt es soweit wie möglich - für den Export)
     * @param aView  muss 0 für Noten-, 1 für Absenzen-Eingaben, 2 für JSModul oder 3 für JSModul/Absenzen sein
     */
    showOverlay: function(aView) {
        if (aView >= 2 && $("td.gradeInput").length == 0) {
            // "Weiter zur Auswertung" entlädt nicht
            return;
        }

        var kursname = $(aView >= 2 ? "td.dialogMainInfo + td" : "span[id$='lblAnlassBezeichnung']").text() || "absent";
        var lines = X.strings[X.lang].views[aView].default_lines.join("\n").replace("%s", kursname) +
            "\n" + X.collectNames(aView, true).join("\n") + "\n";

        if (aView >= 2) {
            $("#tsv-data + div > a:first-child").attr("onClick", 'top.X.acceptOverlay(' + aView + ');').html(X.strings[X.lang].views[aView].accept_button);
        }

        $("#tsv-overlay").show(1000, function() {
            $("textarea", this).focus();
            $("textarea", this).select();
        }).find("textarea").val(lines);
    },

    /**
     * übernimmt die Angaben des Excel-Eingabefelds ins Evento-Formular
     * @param aView  muss 0 für Noten-, 1 für Absenzen-Eingaben, 2 für JSModul oder 3 für JSModul/Absenzen sein
     */
    acceptOverlay: function(aView) {
        var lines = $("#tsv-overlay").hide(1000).find("textarea").val().split("\n");
        var errorColors = {
            "not-found": "#ff6",
            "grade-not-found": "#ff6",
            "name-double": "#fcc",
            "invalid-value": "#fcc",
            "no-number": "#ff6"
        };

        switch (aView) {
            case 0:
                var grades = X.parseGradeData(lines, X.collectNames(aView), X.collectValidGrades(aView));

                // für jede Zeile des Evento-Formulars wird entweder
                // * ein "Name nicht gefunden" Fehler angezeigt, wenn keine Daten verfügbar waren
                // * ein "Ungültiger Wert" Fehler angezeigt, wenn die eingegebene Note in der
                //   Auswahlliste nicht vorkam
                // * der Wert übertragen und kein Fehler angezeigt
                $(".tablelabel + .content1").each(function() {
                    var name = X.trimName($(this).text());
                    var error = [null, null];

                    if (name in grades) {
                        var select = $(this).parent().find("select");
                        if (/^error-(.*)/.test(grades[name])) {
                            error = [RegExp.$1, name];
                        } else if (select.length > 0) // Note aus Auswahlliste auswählen
                        {
                            select.val(select.find("option").filter(function() {
                                return $.trim($(this).text()) == grades[name];
                            }).val() || grades[name]);
                            if ($.trim(select.find("option:selected").text()) != grades[name]) {
                                error = ["invalid-value", grades[name]];
                            }
                        } else // Note frei eingeben
                        {
                            $(this).parent().find(":text").val(grades[name]);
                            if (typeof(grades[name]) != "number") {
                                error = ["no-number", grades[name]];
                            }
                        }
                    } else {
                        error = ["not-found", name];
                    }

                    var errorString = error[0] && X.strings[X.lang].errors[error[0].replace(/-/g, "_")] || "";
                    $(this).parent().children("td").css("background-color", errorColors[error[0]] || "").end()
                        .children("td.errortext").text(errorString.replace("%s", error[1]));
                });
                break;
            case 1:
                var absences = X.parseAbsenceData(lines, X.collectNames(aView));

                // für jede Zeile des Evento-Formulars wird entweder
                // * ein "Name nicht gefunden" Fehler angezeigt, wenn keine Daten verfügbar waren
                // * ein "Keine Zahl" Fehler angezeigt, wenn die eingegebenen Werte keine
                //   gültigen Absenzen-Daten sind
                // * der Wert übertragen und kein Fehler angezeigt
                $("td.tablelabel:first-child, table.WebPart-Adaptive td:first-child").each(function() {
                    var name = X.trimName($(this).text());
                    if (name) {
                        var error = [null, null];

                        if (name in absences) {
                            if (typeof(absences[name]) == "string" && /^error-(.*)/.test(absences[name])) {
                                error = [RegExp.$1, name];
                            } else if (typeof(absences[name][0]) != "number" || typeof(absences[name][1]) != "number") {
                                error = ["invalid-value", "" + absences[name]];
                            } else {
                                $(this).parent().find(":text:eq(0)").val(absences[name][0]);
                                $(this).parent().find(":text:eq(1)").val(absences[name][1]);
                            }
                        } else {
                            error = ["not-found", name];
                        }

                        if ($(this).parent().children(".errortext").length == 0) {
                            $(this).parent().append('<td class="errortext"></td>');
                        }

                        var errorString = error[0] && X.strings[X.lang].errors[error[0].replace(/-/g, "_")] || "";
                        $(this).parent().children("td").css("background-color", errorColors[error[0]] || "").end()
                            .children("td.errortext").text(errorString.replace("%s", error[1]));
                    }
                });
                break;
            case 2:
                var validGrades = X.collectValidGrades(aView);
                var grades = X.parseGradeData(lines, X.collectNames(aView), validGrades);

                // für jede Zeile des Evento-Formulars wird entweder
                // * ein "Name nicht gefunden" Fehler angezeigt, wenn keine Daten verfügbar waren
                // * ein "Ungültiger Wert" Fehler angezeigt, wenn die eingegebene Note in der
                //   Auswahlliste nicht vorkam
                // * der Wert übertragen und kein Fehler angezeigt
                $("td.validationColumn + td").each(function() {
                    var name = X.trimName($(this).text());
                    var error = [null, null];

                    if (name in grades) {
                        var input = $(this).parent().find("td.gradeInput input");
                        if (/^error-(.*)/.test(grades[name])) {
                            error = [RegExp.$1, name];
                        } else {
                            if (validGrades && $.inArray(validGrades, grades[name])) {
                                // bei Zehntelsnoten müssen ganze Werte auf ".0" enden
                                grades[name] = $.grep(validGrades, function(aVal) {
                                    return aVal == grades[name];
                                })[0];
                            }
                            // die Eingabe Server-schonend nur bei Änderung vornehmen
                            if (input.val() != grades[name]) {
                                input.val(grades[name]);
                                // Auto-Speicherung durch Simulation einer Eingabe auslösen
                                input.trigger("keyup").trigger("input").trigger("blur");

                                // Bugfix: manchmal ist die Eingabe beim ersten Versuch nicht
                                // erfolgreich (v.a. Eingabe von Zehntelsnoten)
                                setTimeout(function() {
                                    if (input.val() != grades[name]) {
                                        input.val(grades[name]);
                                        // Auto-Speicherung durch Simulation einer Eingabe auslösen
                                        input.trigger("keyup").trigger("input").trigger("blur");
                                    }
                                }, 500);
                            }

                            if (validGrades && !$.inArray(validGrades, grades[name])) {
                                error = ["invalid-value", grades[name]];
                            } else if (!validGrades && typeof(grades[name]) != "number") {
                                error = ["no-number", grades[name]];
                            }
                        }
                    } else {
                        error = ["not-found", name];
                    }

                    if ($(this).parent().find("span.errortext").length == 0) {
                        $(this).parent().find("td").last().append(" <span class='errortext'></span>");
                    }

                    var errorString = error[0] && X.strings[X.lang].errors[error[0].replace(/-/g, "_")] || "";
                    $(this).parent().children("td").css("background-color", errorColors[error[0]] || "")
                        .children("span.errortext").text(errorString.replace("%s", error[1]));
                });
                break;
            case 3:
                var absences = X.parseAbsenceData(lines, X.collectNames(aView));

                // für jede Zeile des Evento-Formulars wird entweder
                // * ein "Name nicht gefunden" Fehler angezeigt, wenn keine Daten verfügbar waren
                // * ein "Keine Zahl" Fehler angezeigt, wenn die eingegebenen Werte keine
                //   gültigen Absenzen-Daten sind
                // * der Wert übertragen und kein Fehler angezeigt
                $("td.validationColumn + td").each(function() {
                    var name = X.trimName($(this).text());
                    var error = [null, null];

                    if (name in absences) {
                        if (typeof(absences[name]) == "string" && /^error-(.*)/.test(absences[name])) {
                            error = [RegExp.$1, name];
                        } else if (typeof(absences[name][0]) != "number" || typeof(absences[name][1]) != "number") {
                            error = ["invalid-value", "" + absences[name]];
                        } else {
                            for (var i = 0; i < 2; i++) {
                                var el = $(this).parent().find("td:not(.gradeInput) input[type=text]").eq(-2 + i);
                                // die Eingabe Server-schonend nur bei Änderung vornehmen
                                if ((el.val() || 0) != absences[name][i]) {
                                    el.val(absences[name][i]);
                                    // Auto-Speicherung durch Simulation einer Eingabe auslösen
                                    el.trigger("keyup").trigger("input").trigger("blur");
                                }
                            }
                        }
                    } else {
                        error = ["not-found", name];
                    }

                    if ($(this).parent().find("span.errortext").length == 0) {
                        $(this).parent().find("td").last().append(" <span class='errortext'></span>");
                    }

                    var errorString = error[0] && X.strings[X.lang].errors[error[0].replace(/-/g, "_")] || "";
                    $(this).parent().children("td").css("background-color", errorColors[error[0]] || "")
                        .children("span.errortext").text(errorString.replace("%s", error[1]));
                });
                break;
        }
    },

    /**
     * bricht die Excel-Eingabe ab
     */
    cancelOverlay: function() {
        $("#tsv-overlay").hide(1000);
    },

    /**
     * @param aView  muss 0 für Noten-, 1 für Absenzen-Eingaben, 2 für JSModul oder 3 für JSModul/Absenzen sein
     * @param aIncData  gibt an, ob die Noten/Absenzen zu den SchülerInnennamen
     *                  hinzugefügt werden sollen (mit Tabulatoren getrennt)
     * @returns die Namen sämtlicher SchülerInnen aus dem Evento-Formular
     *          (optional inklusive bereits eingegebener Noten/Absenzen)
     */
    collectNames: function(aView, aIncData) {
        var values = [];

        var nameCell = aView >= 2 ? "td.validationColumn + td" :
            aView == 0 ? ".tablelabel + .content1" :
            "td.tablelabel:first-child, table.WebPart-Adaptive td:first-child";
        $(nameCell).each(function() {
            var name = X.trimName($(this).text());
            if (name) {
                var data = "";
                if (aIncData) {
                    switch (aView) {
                        case 0:
                            var select = $(this).parent().find("select");
                            if (select.length > 0) // Note aus fester Auswahl
                            {
                                data = $.trim($("option:selected", select).text() || "");
                            } else // Note aus freier Eingabe
                            {
                                data = $.trim($(this).parent().find(":text").val());
                            }
                            break;
                        case 1:
                            // Absenzen aus zwei Textfeldern sammeln
                            data = [];
                            $(this).parent().find(":text").each(function() {
                                data.push($.trim($(this).val()).replace(/\.0+$/, ""));
                            });
                            data = data.concat(["", ""]).slice(0, 2);
                            data = data.join("") ? data.join("\t") : "";
                            break;
                        case 2:
                            data = $.trim($(this).parent().find("td.gradeInput input").val());
                            break;
                        case 3:
                            // Absenzen aus zwei Textfeldern sammeln
                            data = [];
                            $(this).parent().find("td:not(.gradeInput) input[type=text]").slice(-2).each(function() {
                                data.push($.trim($(this).val()).replace(/\.0+$/, ""));
                            });
                            data = data.concat(["", ""]).slice(0, 2);
                            data = data.join("") ? data.join("\t") : "";
                            break;
                    }
                }

                values.push(name + (data ? "\t" + data : ""));
            }
        });

        return values;
    },

    /**
     * @param aView  muss 0 für Noten-, 1 für Absenzen-Eingaben oder 2 für JSModul sein
     * @returns eine Liste sämtlicher gültigen Notenwerte aus der Auswahlliste
     *          oder |null| falls die Notenwerte in ein Textfeld eingegeben werden
     */
    collectValidGrades: function(aView) {
        var firstSelect = $(aView == 2 ? "td.gradeInput select" : ".tablelabel + .content1").parent().find("select").get(0);
        // JSModul verwendet seit 2019 eine Liste anstelle einer Auswahl
        if (!firstSelect && aView == 2 && $("td.gradeInput ul.dialogContextMenu").length > 0) {
            firstSelect = {
                options: $("td.gradeInput ul.dialogContextMenu").first().find("li")
            };
        }
        if (!firstSelect) {
            return null;
        }

        var values = [];

        $.each(firstSelect.options, function() {
            var value = $.trim($(this).text());
            // JSModule verwendet "<>", wenn nichts ausgewählt ist
            if (value && (aView != 2 || value != "<>")) {
                values.push(value);
            }
        });

        return values;
    },

    /**
     * @returns welche Ansicht das Dokument bietet (0: Noteneingabe, 1: Absenzeneingabe, 2: JSModul, -1: nicht unterstützt)
     */
    viewType: function() {
        if ($("form[action*='Brn_QualifikationDurchDozenten.aspx']").length > 0) {
            return 0;
        }

        // Kleinschreibung für cst_pages (siehe EVT CR-11450)
        if ($("form[action*='fct=AnmeldungMultiSave'], form[action*='Brn_Absenzverwaltung_ProAnlass.aspx']" +
                ", form[action*='fct=anmeldungmultisave'], form[action*='brn_absenzverwaltung_proanlass.aspx']").length > 0) {
            return 1;
        }
        if ($("form[action*='./brn_qualifikationdurchdozenten.aspx']").length > 0) {
            return 2;
        }
        return -1;
    },

    /**
     * Das akzeptierte Datenformat sind Tabulator-getrennte Werte, wobei die erste
     * Spalte die Namen im Format "Nachname Vorname" oder "Vorname Nachname" enthalten
     * muss oder aber die ersten zwei Spalten Nach- und Vornamen (in beliebiger, aber
     * konsistenter Reihenfolge) enthalten müssen.
     * 
     * @param aData  Daten, aus welchen die Namen der SchülerInnen und die weiteren Daten
     *               bestimmt werden sollen
     * @param aKnownNames  eine Liste der dem System bekannten Namen
     * @param aValidator  eine optionale Funktion, welche bestimmt, ob es sich bei einem
     *                    Zellwert um einen gültigen Wert handelt; die zurückgegebene
     *                    Liste enthält nur Daten bis zur letzten Spalte mit gültigen Werten
     * @returns eine Liste von Listen, deren erstes Element jeweils ein normierter Name ist
     */
    parseDataHelper: function(aData, aKnownNames, aValidator) {
        var lessFancy = {};
        var knownUnfancy = [];
        $.each(aKnownNames, function() {
            var name = X.unfancyName(this);
            lessFancy[name] = name in lessFancy ? null : this;
            knownUnfancy.push(name);
        });

        // zuerst muss das Muster bestimmt werden, in welchem Namen und Werte auftreten;
        // das meist-verwendete Namensschema und die letzte Spalte mit Zahlen werden verwendet
        var stats = {
            normal: 0,
            split: 0,
            split_rev: 0,
            rotate: 0,
            gradeRow: 1
        };
        for (var i = 0; i < aData.length; i++) {
            // ignoriere Leerzeilen und Kommentarzeilen
            if (!aData[i] || aData[i].charAt(0) == "#") {
                aData[i] = null;
                continue;
            }

            // Bugfix: split mit RegExp funktioniert im MSIE nicht zuverlässig
            aData[i] = $.map(aData[i].split("\t"), $.trim);

            stats.normal += X.resolveName(X.unfancyName(aData[i][0]), knownUnfancy) in lessFancy ? 1 : 0;
            stats.split += X.resolveName(X.unfancyName(aData[i].slice(0, 2).join(" ")), knownUnfancy) in lessFancy ? 1 : 0;
            stats.split_rev += X.resolveName(X.unfancyName(aData[i].slice(0, 2).reverse().join(" ")), knownUnfancy) in lessFancy ? 1 : 0;

            $.each(X.rotateName(X.unfancyName(aData[i][0])), function() {
                if (X.resolveName(this, knownUnfancy) in lessFancy) {
                    stats.rotate++;
                    return false;
                }
            });

            $.each(aData[i], function(aRow) {
                if (aRow > stats.gradeRow && (aValidator ? aValidator(this) : this)) {
                    stats.gradeRow = aRow;
                }
            });
        }

        // Anzahl Zellen, die jede Zeile mindestens haben muss
        var padding = [];
        for (i = 0; i < stats.gradeRow; i++) {
            padding.push("");
        }

        var parsedLines = [];
        for (i = 0; i < aData.length; i++) {
            if (!aData[i]) {
                continue;
            }
            aData[i] = aData[i].concat(padding).slice(0, stats.gradeRow + 1);
            if (stats.split > Math.max(stats.normal, stats.split_rev, stats.rotate)) {
                var name = aData[i].splice(0, 2).join(" ");
            } else if (stats.split_rev > Math.max(stats.normal, stats.split, stats.rotate)) {
                name = aData[i].splice(0, 2).reverse().join(" ");
            } else if (stats.rotate > Math.max(stats.normal, stats.split, stats.split_rev)) {
                $.each(X.rotateName(X.unfancyName(aData[i].splice(0, 1)[0])), function() {
                    name = this;
                    if (X.resolveName(this, knownUnfancy) in lessFancy) {
                        return false;
                    }
                });
            } else {
                name = aData[i].splice(0, 1)[0];
            }
            name = lessFancy[X.resolveName(X.unfancyName(name), knownUnfancy)] || name;

            parsedLines.push([name].concat(aData[i]));
        }

        return parsedLines;
    },

    /**
     * liest von Excel kopierte Daten nach dem allgemeinen Schema von parseDataHelper ein,
     * wobei die die auf die Namen folgenden Daten mindestens eine Noten-Spalte enthalten
     * sollten
     * 
     * Gültig sind z.B. die folgenden Eingaben (Hinweis: die einzelnen Zeilen sind
     * gültig, nicht aber die ganze Tabelle, da die Namen nicht in konsistenter
     * Form vorliegen):
     * 
     * Name	Vorname		6
     * Name Vorname		5.5
     * Nom	Prénom	<beliebiger Inhalt>	5.0
     * Given Name	Family Name	3.25	4.5
     * Cognome Nome	-1	4.00
     * Nombre Apellido		3 1/2
     * 
     * Die oben ausgelesenen Notenwerte sind 6, 5.5, 5, 4.5, 4 und 3.5.
     * 
     * @param aData  Daten, aus welchen die Namen der SchülerInnen und ihre Noten
     *               bestimmt werden sollen
     * @param aKnownNames  eine Liste der dem System bekannten Namen
     * @param aValidGrades  eine Liste der vom System akzeptierten Noten
     * @returns einen Hash, welcher jedem/r SchülerIn eine Note zuweist
     */
    parseGradeData: function(aData, aKnownNames, aValidGrades) {
        function validate(aValue) {
            var value = X.parseNumber(aValue);
            return !isNaN(value) || aValidGrades && $.inArray(value, aValidGrades) > -1;
        }
        var lines = X.parseDataHelper(aData, aKnownNames, validate);

        var grades = {};
        for (i = 0; i < lines.length; i++) {
            var name = lines[i][0];
            var grade = lines[i].slice(-1)[0];

            grades[name] = name in grades ? "error-name-double" : X.parseNumber(grade || "") || grade || "error-grade-not-found";
        }

        return grades;
    },

    /**
     * liest von Excel kopierte Daten nach dem allgemeinen Schema von parseDataHelper ein,
     * wobei die die auf die Namen folgenden Daten mindestens zwei Absenzen-Spalte enthalten
     * sollten
     * 
     * Gültig sind z.B. die folgenden Eingaben (Hinweis: die einzelnen Zeilen sind
     * gültig, nicht aber die ganze Tabelle, da die Namen nicht in konsistenter
     * Form vorliegen):
     * 
     * Name	Vorname		2	0
     * Name Vorname		3	1
     * Nom	Prénom	<beliebiger Inhalt>	0	5
     * Given Name	Family Name	-1		6
     * Cognome Nome	Vermerk		
     * Nombre Apellido		0	0
     * 
     * Die oben ausgelesenen Absenzen sind (2, 0), (3, 1), (0, 5), (0, 6), (0, 0), (0, 0).
     * 
     * @param aData  Daten, aus welchen die Namen der SchülerInnen und ihre Absenzen
     *               bestimmt werden sollen
     * @param aKnownNames  eine Liste der dem System bekannten Namen
     * @returns einen Hash, welcher jedem/r SchülerIn zwei Absenzen-Zahlen zuweist
     */
    parseAbsenceData: function(aData, aKnownNames) {
        function validate(aValue) {
            return /^(\d{0,3})(?:\.0+)?$/.test(aValue);
        }
        var lines = X.parseDataHelper(aData, aKnownNames, validate);

        var absences = {};
        for (i = 0; i < lines.length; i++) {
            var name = lines[i][0];
            var absence = $.map(lines[i].slice(1).slice(-2), function(aValue) {
                return validate(aValue || "0") ? parseInt(RegExp.$1) : aValue;
            });
            if (!absence[1]) {
                absence[1] = 0;
            }

            absences[name] = name in absences ? "error-name-double" : absence;
        }

        return absences;
    },

    /**
     * @param aString  möglicherweise eine Zahl (Dezimalbruch oder gemeiner Bruch)
     * @returns die Zahl als Zahl oder NaN
     */
    parseNumber: function(aString) {
        if (/^[1-6](?:\.\d+)?$/.test(aString)) // Dezimalbruch
        {
            return parseFloat(aString);
        }
        if (/[1-6],\d+$/.test(aString)) // Dezimalbruch mit Komma
        {
            return parseFloat(aString.replace(",", "."));
        }
        if (/^([1-6]) (\d+)\/(\d+)$/.test(aString) && RegExp.$3 != 0 && RegExp.$3 - RegExp.$2 > 0) // gemeiner Bruch
        {
            return parseInt(RegExp.$1) + parseInt(RegExp.$2) / parseInt(RegExp.$3);
        }
        return NaN;
    },

    /**
     * @param aName  ein Name aus dem Evento-Formular
     * @returns den Namen mit normalisierten Leerzeichen
     */
    trimName: function(aName) {
        // Bugfix: MSIE produziert unter Umständen ein geschütztes Leerzeichen (&#160;)
        return $.trim(aName.replace(/[\s\xA0]+/g, " "));
    },

    /**
     * @param aName  ein Name
     * @returns eine normiertere Version dieses Namens ohne Umlaute,
     *          geläufige Akzente, Bindestriche und Gross-/Kleinschreibung
     */
    unfancyName: function(aName) {
        var lessFancy = {
            "äÄ": "ae",
            "öÖ": "oe",
            "üÜ": "ue",
            "àÀáÁâÂ": "a",
            "éÉèÈëËêÊ": "e",
            "ïÏíÍîÎ": "i",
            "óÓôÔ": "o",
            "úÚûÛ": "u",
            "ñÑ": "n",
            "-": " "
        };

        for (var fancy in lessFancy) {
            aName = aName.replace(new RegExp("[" + fancy + "]", "g"), lessFancy[fancy]);
        }

        return $.trim(aName.toLowerCase().replace(/\s+/g, " "));
    },

    /**
     * @param aName  ein Name
     * @returns eine Liste mit dem Namen derart umgestellt, dass der erste
     *          Nachname an zweiter, dritter, etc. Stelle steht
     */
    rotateName: function(aName) {
        var variants = [];

        var parts = aName.split(" ");
        for (var i = 1; i < parts.length; i++) {
            variants.push(parts.slice(i).concat(parts.slice(0, i)).join(" "));
        }

        return variants;
    },

    /**
     * Sucht den Namen aName in aKnownNames, wobei aName nicht sämtliche
     * Namensteile enthalten muss (aber mindestens den ersten Nachnamen).
     * 
     * @param aName  ein Name, beginnend mit dem ersten Nachnamen
     * @param aKnownNames  eine Liste von dem System bekannten Namen
     * @returns den Namen, wie er in aKnownNames auftritt, oder |null|
     */
    resolveName: function(aName, aKnownNames) {
        var found = null;

        var nameParts = aName.split(" ");
        for (var i = 0; i < aKnownNames.length; i++) {
            if (X.matchNameParts(nameParts, aKnownNames[i].split(" "))) {
                if (!found) {
                    found = aKnownNames[i];
                } else {
                    // Name ist nicht eindeutig zuweisbar
                    return null;
                }
            }
        }

        return found;
    },

    /**
     * Prüft, ob sämtliche Teile von aParts in der exakten Reihenfolge auch
     * in aFull auftreten. Der erste Teil muss übereinstimmen (1. Nachname)
     * und mindestens zwei Namensteile müssen vorkommen.
     * 
     * @param aParts  eine Liste von (Namens)teilen
     * @param aFull  eine Liste von (Namens)teilen
     * @returns ob sämtliche Teile von aParts in aFull auftreten
     */
    matchNameParts: function(aParts, aFull) {
        if (aParts.length < Math.min(2, aFull.length) || aParts[0] != aFull[0]) {
            return false;
        }

        var ix = 0;

        $.each(aFull, function() {
            if (this == aParts[ix]) {
                ix++;
            }
        });

        return ix == aParts.length;
    },

    /**
     * @param aFunc  die zu ersetzende Nebeneffekt-freie Funktion
     * @returns eine memoisierte Version dieser Funktion
     */
    memoizeFunction: function(aFunc) {
        var cache = {};

        return function() {
            if (arguments.length != 1 || typeof(arguments[0]) != "string") {
                return aFunc.apply(this, arguments);
            }
            if (!(arguments[0] in cache)) {
                return (cache[arguments[0]] = aFunc.apply(this, arguments));
            }
            return cache[arguments[0]];
        };
    },

    /**
     * testet den Datenparser; muss in der Konsole manuell aufgerufen werden:
     * 
     * X.test()
     */
    test: function() {
        var tests = [
            [ // Muster: Nachname<Tab>Vorname<Tab>Note
                "Name	Vorname	1",
                "Nom	Prénom	3.5",
                "Family Name	Given Name	2.2",
                "Appellido	Nombre	3.5", // Schreibfehler
                "Name	Unbekannt	2.5", // unbekannter Name
                "Nàlizätión	Iñtërnâtiô	6",
                "Cognome	Nome", // fehlende Note
                "Apellido Nombre	5.5", // Leerschlag statt Tabulator
                "Sportler	Profi	disp",
                "Tester	Beta	Besucht" // Gross-/Kleinschreibung des Prädikats
            ],
            [ // Muster: Nachname<Leerschlag>Vorname<Tab>beliebig<Tab>Note als gemeiner Bruch
                "Name Vorname	1.2	1",
                "Nom Prénom	3.3	3 1/2",
                "Family Name Given Name	++	2 2/10",
                "Nàlizätión Iñtërnâtiô		6",
                "Apellido Nombre	4 1/5", // Note in falscher Spalte
                "Cognome	Nome	2.0", // Tabulator statt Leerschlag
                "Apellido Nombre		4 0/2", // ungültige Note
                "Sportler Profi		disp",
                "Tester Beta		besucht nicht" // unbekanntes Prädikat
            ],
            [ // Muster: Vorname<Tab>Nachname<Tab>beliebig<Tab>Note auf zwei Dezimalstellen
                "Vorname	Name	/!\	1.00",
                "Given	Family	/!\	2.20", // jeweils nur der erste Namensteil
                "Prenom	Nom	/!\	3.50",
                "Internatio	Nalizaetion	#	6.00",
                "Nome	Cognome		4 .5", // ungültige Note
                "Nombre	Apellido	?	7.5", // ungültige Note
                "Profi	Sportler		disp"
            ],
            [ // Muster: VORNAME<Leerschlag>NAME<Tab>NOTE<Tab>beliebig
                "VORNAME NAME	1	1+",
                "PRÉNOM NOM	3 1/2",
                "GIVEN NAME FAMILY NAME	2 1/5",
                "IÑTËRNÂTIÔ NÀLIZÄTIÓN	6",
                "NOME COGNOME	2 1.3/2	NaN", // ungültige Note
                "NOMBRE APELLIDO	4", // zweimal der-
                "nómbré ápéllídó	5", // selbe Name
                "PROFI SPORTLER	disp",
                "Beta Tester	xbesucht" // Tippfehler im Prädikat
            ]
        ];
        var absenceTests = [
            [ // Muster: Nachname<Tab>Vorname<Tab>beliebig<Tab>Entschuldigte<Tab>Unentschuldigte
                "Name	Vorname",
                "Nom	Prénom	x	2	2",
                "Family Name	Given Name		1	1",
                "Appellido	Nombre		0	0", // Schreibfehler
                "Name	Unbekannt		1	1", // unbekannter Name
                "Nàlizätión	Iñtërnâtiô		3.0	3.00",
                "Cognome	Nome	x	?	?", // keine gültigen Werte
                "Apellido Nombre		4	4" // Leerschlag statt Tabulator
            ]
        ];

        var knownNames = [
            "Name Vorname",
            "Family Name Given Name",
            "Nom Prénom",
            "Cognome Nome",
            "Apellido Nombre",
            "Nàlizätión Iñtërnâtiô",
            "Sportler Profi",
            "Tester Beta"
        ];
        var validGrades = "1 2.2 3.5 5.5 6 disp besucht".split(" ");
        var output = {
            "Name Vorname": 1,
            "Family Name Given Name": 2.2,
            "Nom Prénom": 3.5,
            "Nàlizätión Iñtërnâtiô": 6,
            "Sportler Profi": "disp"
        };
        var absenceOutput = {
            "Name Vorname": 0,
            "Family Name Given Name": 1,
            "Nom Prénom": 2,
            "Nàlizätión Iñtërnâtiô": 3,
            "Cognome Nome": "?"
        };

        var errors = [];

        function assert(aCondition, aError) {
            if (!aCondition) {
                errors.push(aError);
            }
        }

        $.each(tests, function(i) {
            var result = X.parseGradeData(this, knownNames, validGrades);
            var count = 0;
            for (var name in output) {
                assert(name in result, "Test " + (i + 1) + ": '" + name + "' nicht gefunden");
                if (name in result) {
                    assert(result[name] === output[name], "Test " + (i + 1) + " für '" + name + "': " + result[name] + " != " + output[name]);
                }
                count++;
            }
            for (name in result) {
                if ($.inArray(name, knownNames) > -1 && $.inArray(result[name].toString(), validGrades) > -1) {
                    count--;
                }
            }
            assert(count >= 0, "Test " + (i + 1) + ": Differenz von " + -count + " zur Anzahl erwarteter Ergebnisse");
        });

        $.each(absenceTests, function(i) {
            var result = X.parseAbsenceData(this, knownNames);
            var count = 0;
            for (var name in absenceOutput) {
                assert(name in result, "Absenzen-Test " + (i + 1) + ": '" + name + "' nicht gefunden");
                if (name in result) {
                    assert(result[name][0] === absenceOutput[name] && result[name][1] === absenceOutput[name], "Absenzen-Test " + (i + 1) + " für '" + name + "': " + [result[name], result[name]] + " != " + absenceOutput[name]);
                }
                count++;
            }
            for (name in result) {
                if ($.inArray(name, knownNames) > -1 && typeof(result[name]) != "string" && result[name][0] === result[name][1]) {
                    count--;
                }
            }
            assert(count >= 0, "Absenzen-Test " + (i + 1) + ": Differenz von " + -count + " zur Anzahl erwarteter Ergebnisse");
        });

        return errors.join("\n") || "Tests bestanden.";
    }
};

X.init();