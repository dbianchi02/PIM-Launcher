# PIM Launcher

GUI in PowerShell per attivare i ruoli **Microsoft Entra ID PIM** sui tenant dei clienti e aprire subito i portali di amministrazione con una sessione browser che contiene il ruolo appena attivato.

<!-- ![Screenshot](docs/screenshot.png) -->

## Perché

Dopo l'attivazione di un ruolo in PIM, i token già emessi non contengono il nuovo ruolo: per usarlo serve un token nuovo, quindi sign-out, refresh della sessione o finestra in incognito. Chi gestisce più tenant e più portali (Entra, Defender, Purview, Exchange, Teams, SharePoint...) ripete questa operazione decine di volte a settimana.

PIM Launcher unisce i passaggi: attivi i ruoli, aspetta che risultino attivi e apre i portali del cliente in una sessione pulita.

> Non è un bypass. Il comportamento dei token è voluto da Microsoft e non è configurabile: lo strumento rende solo più rapido ottenere un token nuovo.

## Cosa fa

- Si connette al tenant scelto con Microsoft Graph e mostra i **ruoli Entra eleggibili** (con stato attivo e scadenza).
- **Attiva più ruoli insieme**, con giustificazione, durata e numero di ticket (opzionale).
- Attende che i ruoli risultino attivi.
- Apre i **portali del cliente** in Microsoft Edge, tutti insieme o uno alla volta.

## Requisiti

- Windows (l'interfaccia usa WPF)
- PowerShell 7
- Microsoft Edge
- Modulo `Microsoft.Graph.Authentication`

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

## Avvio

1. Copia `tenants.example.json` in `tenants.json` (stessa cartella dello script) e compilalo.
2. Da una console PowerShell 7:

```powershell
Unblock-File .\PIM-Launcher.ps1, .\tenants.json
.\PIM-Launcher.ps1
```

3. Scegli il cliente, premi **Connetti**, spunta i ruoli, inserisci la giustificazione e premi **Attiva selezionati**.

> `tenants.json` contiene i dati dei tuoi clienti ed è escluso dalla repository tramite `.gitignore`. Non committarlo.

<img width="856" height="640" alt="image" src="https://github.com/user-attachments/assets/a5ec1389-0bfa-425d-b5c4-ed0899e862af" />

## Configurazione

`tenants.json` ha una parte globale e l'elenco dei clienti.

**Globale**

| Campo | Descrizione |
|---|---|
| `ClientId` | Opzionale. ID di una tua app registration (client pubblico). Se vuoto usa l'app *Microsoft Graph Command Line Tools* |
| `DefaultJustification` | Giustificazione precompilata |
| `DefaultDurationHours` | Durata predefinita (1, 2, 4 o 8 ore) |
| `Portals` | Elenco dei portali disponibili, con URL modello |

Gli URL dei portali possono contenere questi segnaposto, sostituiti con i dati del cliente:

| Segnaposto | Valore |
|---|---|
| `{tenantId}` | `TenantId` del cliente |
| `{upn}` | `Upn` del cliente |
| `{spo}` | `Spo` del cliente (nome del tenant SharePoint, senza `-admin`) |

**Per ogni cliente (`Tenants`)**

| Campo | Obbligatorio | Descrizione |
|---|---|---|
| `Name` | Sì | Nome mostrato nel menu |
| `TenantId` | Sì | ID del tenant |
| `Upn` | Solo se un portale usa `{upn}` | Account con cui accedi |
| `Spo` | Solo per il portale SharePoint | Per `contoso-admin.sharepoint.com` scrivi `contoso` |
| `EdgeProfile` | Solo con la modalità profilo | Nome della **cartella** del profilo Edge (`Default`, `Profile 1`...) |
| `PortalNames` | No | Portali da aprire per quel cliente. Se omesso, tutti |

### Trovare la cartella del profilo Edge

I profili sono in `%LOCALAPPDATA%\Microsoft\Edge\User Data`. Per elencarli con l'account associato si può utilizzare PowerShell:

```powershell
$ls = Get-Content "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Local State" -Raw | ConvertFrom-Json -AsHashtable
$ls.profile.info_cache.GetEnumerator() | ForEach-Object { '{0}  ->  {1}  ({2})' -f $_.Key, $_.Value.name, $_.Value.user_name }
```

## Modalità browser

| Modalità | Comportamento |
|---|---|
| **Profilo Edge del cliente** (predefinita) | Apre Edge nel profilo indicato in `EdgeProfile`. Comodo per l'SSO, ma il portale può riusare un token vecchio: in quel caso serve un sign-out manuale |
| **Isolata e nuova** | Edge parte con un profilo temporaneo vuoto in `%TEMP%\PIMLauncher` (rimosso dopo un giorno). Token sempre nuovo, login da rifare ogni volta |
| **InPrivate** | Finestra InPrivate. Le finestre InPrivate condividono la stessa sessione: chiudi le precedenti prima di cambiare cliente |

Lo script usa sempre Edge, non il browser predefinito di sistema.

## Permessi Microsoft Graph

Scope delegati richiesti: `RoleManagement.ReadWrite.Directory` e `User.Read`.

In un tenant cliente potrebbe servire il **consenso amministratore** per l'app *Microsoft Graph Command Line Tools*, oppure puoi registrare una tua app e indicarne l'ID in `ClientId`.

## Limiti noti

- Gestisce solo i **ruoli Entra ID**: non ruoli delle risorse Azure né PIM for Groups.
- **Step-up non gestito**: se la policy del ruolo richiede un authentication context o un MFA fresco, l'attivazione fallisce e il log lo segnala.
- I ruoli che richiedono **approvazione** vengono inviati e segnati come in attesa, senza altro.
- Exchange, SharePoint e Teams possono impiegare diversi minuti a propagare il ruolo: nessun token nuovo lo accelera.
- Alcuni portali non accettano parametri di tenant o utente nell'URL: modifica `Portals` nel JSON se un link non si comporta come previsto.
- Solo Windows e solo Edge.

## Sicurezza e privacy

- Lo script non salva né invia credenziali o token: l'autenticazione è interamente gestita dal modulo Microsoft Graph.
- Le attivazioni sono richieste PIM normali e compaiono nei log di audit di Entra ID come qualsiasi attivazione dal portale.
- Lo script comunica solo con Microsoft Graph e con gli endpoint di autenticazione Microsoft. Nessuna telemetria e nessun servizio di terze parti.

## Licenza

[MIT](LICENSE)
