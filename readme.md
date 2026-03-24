# Canon X-07 Serial Fast Loader (GUI)

![Screenshot](./screenshot.png)

## 🇫🇷 Français

### Présentation

Cet outil, basé sur une interface graphique **Tkinter**, permet de transférer des programmes vers un **Canon X-07** via une liaison **série**.

Il prend en charge :

* **BASIC (listing texte)** : envoi d’un fichier `.txt` / `.bas` en le « tapant » ligne par ligne
* **BASIC (cassette en flux brut)** : envoi *et* réception de fichiers `.cas` / `.k7` via `LOAD"COM:"` et `SAVE"COM:"`
* **ASM (rapide)** : transfert d’un loader puis envoi rapide d’un binaire `.bin`
* **Clavier distant (REMOTE KEYBOARD)** : saisie depuis le PC vers le X-07

L’application fournit :

* une console de logs horodatée `[hh:mm:ss]`
* une barre de progression
* un bouton d’annulation
* des boutons de touches spéciales

---

### Prérequis

* Python 3.10 ou plus récent
* `pyserial` (installé automatiquement si nécessaire)

---

### Matériel requis

* Canon X-07 fonctionnel
* Adaptateur USB ↔ série (FT232, CP2102, CH340…)

Optionnel :

* câble avec **RTS/CTS**

⚠️ Particularité du Canon X-07 :
les signaux **RX, TX, RTS et CTS sont inversés**.

Exemple avec FT232RL (FT_Prog) :

```
Invert TXD : ON
Invert RXD : ON
Invert RTS : ON
Invert CTS : ON
```

⚠️ Sur certains montages, ajouter des **pull-down (~10 kΩ)** sur :

* RX
* CTS

Ces résistances permettent de fixer un **état logique stable** lorsque les lignes ne sont pas activement pilotées par le Canon X-07.

---

### Lancer l’application

```bash
python x07_loader.pyw
```

1. Sélectionner le **port COM** dans *Serial settings*.
2. Ajuster si besoin :

* **Typing baud (8N2)** : vitesse utilisée pour la saisie BASIC, le clavier distant et les flux cassette.
* **Loader baud (8N2)** : vitesse utilisée pour le transfert rapide ASM.
* **CHAR(s)** / **LINE(s)** : délais (en secondes) entre caractères / entre lignes lors de la saisie BASIC.
* **RTS/CTS cable** : active le **contrôle de flux matériel RTS/CTS** si votre câble supporte ces lignes.

---

### ⚠️ Mode « Slave » du Canon X-07 (requis)

À activer pour :

* **BASIC texte** (`.txt` / `.bas`)
* **ASM** (One click : fast loader + `.bin`)
* **REMOTE KEYBOARD**

Dans ce mode :

* le périphérique série (`COM:`) devient l’**entrée console**,
* le clavier local du X-07 est ignoré,
* les caractères reçus sont interprétés comme des frappes clavier.

#### Entrer en mode slave

À saisir sur le Canon X-07 :

```basic
INIT#5,"COM:
EXEC&HEE1F
```

#### Quitter le mode slave

Deux méthodes :

1. Redémarrer le Canon X-07 (OFF puis ON).
2. Envoyer depuis le port série :

```basic
EXEC&HEE33
```

Le bouton **Disable slave (EXEC&HEE33)** de l’application permet de quitter ce mode à distance.

---

# BASIC

### 1) BASIC texte (.txt / .bas) via SLAVE mode

Objectif : transférer un listing BASIC texte en le « tapant » comme si les lignes étaient saisies au clavier.

Procédure :

1. Mettre le X-07 en **SLAVE mode**.
2. Dans la section **Text listing (.txt/.bas)** :

   * **Select .txt/.bas…**
   * **Send BASIC**

Fonctionnement :

* l’outil envoie les lignes en **8N2**, avec les délais **CHAR(s)** et **LINE(s)**,
* à la fin, l’outil envoie `EXEC&HEE33` pour relâcher l’état console distant (si actif).

---

### 2) BASIC cassette stream (.cas / .k7) via LOAD "COM:" / SAVE "COM:" (raw bytes)

Objectif : échanger un programme “cassette” sous forme de flux d’octets via le port série, sans mode *slave*.

#### a) Envoyer un flux (PC → X-07) avec `LOAD"COM:"`

Côté Canon X-07 :

1. **Ne pas** être en mode slave.
2. Lancer la commande :

```basic
LOAD"COM:"
```

3. Appuyer sur **RETURN** (le X-07 attend alors le flux d’octets).

Côté PC :

1. Dans la section **Cassette stream (.cas/.k7)** :

   * **Select .cas/.k7…**
   * (optionnel) **Inspect header**
   * **Send raw (LOAD "COM:")**

Notes importantes :

* L’envoi est réalisé en **8N2**.
* Le point de départ d’envoi est **fixe** : le transfert commence à l’offset **`0x0010`** du fichier (base validée).
* Le bouton **Inspect header** affiche uniquement un **preview @0x0000** (aide au diagnostic / comparaison de fichiers).

#### b) Recevoir un flux (X-07 → PC) avec `SAVE"COM:"`

Côté Canon X-07 :

1. **Ne pas** être en mode slave.
2. Lancer la commande :

```basic
SAVE"COM:"
```

(ou `SAVE"COM:nom"` selon vos habitudes), puis **RETURN**.

Côté PC :

1. Dans la section **Cassette stream (.cas/.k7)** :

   * **Receive raw (SAVE "COM:")**
   * choisir le fichier de destination `.cas`

Fonctionnement :

* la capture se fait en **8N2**,
* l’enregistrement se termine automatiquement après une courte période d’inactivité (timeout),
* le fichier `.cas` enregistré côté PC contient :

  * un **header 16 octets** (10×`D3` + **nom sur 6 caractères** issu du nom de fichier, tronqué/paddé),
  * suivi du **flux brut** reçu depuis le X-07.

---

# ASM (fast loader via COM)

Objectif : transférer rapidement un binaire assembleur `.bin` via un loader optimisé.

Principe :

1. Un loader est envoyé via `LOAD"COM:"` (flux cassette).
2. Ce loader configure la liaison série à haute vitesse (ex : 8000 bauds).
3. Le PC envoie ensuite le binaire sous forme ASCII hex.
4. Le loader copie les données en mémoire puis exécute automatiquement le programme.

---

## Méthodes

### Send ASM loader

Permet d’envoyer uniquement le loader.

Sur le X-07 :

```basic
LOAD"COM:"
```

Puis sur PC :

* cliquer sur **Send ASM loader**

Le loader est transféré comme un flux cassette (raw).

---

### Send ASM (loader running)

Permet d’envoyer uniquement le binaire, une fois le loader actif.

* Le transfert utilise le baud **Loader baud (8N2)**
* nécessite que le loader soit actif (`RUN`)
* Aucun délai logiciel n’est utilisé (transfert rapide)

⚠️ Important :

Le script envoie automatiquement un préfixe de synchronisation (primer) pour fiabiliser le transfert.

---

### One click: loader + ASM

Automatique :

1. `LOAD"COM:"`
2. envoi loader
3. `EXEC&HEE33:RUN`
4. envoi binaire

---

## Adresses

* **Loader addr** : adresse du loader (ex : `0x1800`)
* **ASM addr** : adresse du programme (ex : `0x2000`)

⚠️ Doit correspondre à :

```asm
ORG $2000
```

## Adresse de chargement

### Loader addr

Le champ **Loader addr** définit l’adresse mémoire où le loader est copié (par défaut : `0x1800`).

### ASM addr

Le champ **ASM addr** définit l’adresse mémoire où le binaire `.bin` est copié (par défaut : `0x2000`).

⚠️ Cette adresse doit correspondre à celle utilisée à l’assemblage.

Exemple :

```asm
ORG $2000
```

Si le binaire est copié à une autre adresse que celle prévue (ORG), alors :

* les `JP` / `CALL`,
* les accès mémoire,
* et les données référencées

peuvent devenir invalides (comportement imprévisible ou crash).

---

# Remote keyboard via SLAVE mode

Objectif : contrôler le X-07 depuis le PC comme un « terminal » (frappes + touches spéciales).

Pré-requis :

1. Mettre le X-07 en **SLAVE mode**.
2. Activer **REMOTE KEYBOARD: ON**.

Fonctionnement :

* Quand **REMOTE KEYBOARD est ON**, l’application ouvre une session série persistante en **8N2**.
* La zone de saisie envoie directement les caractères ASCII imprimables.
* Les boutons envoient des **codes de touches** (HOME/CLR/INS/DEL/BREAK/flèches/etc.) conformes au X-07.

Sécurité :

* Pendant un transfert (BASIC/ASM/CAS), le clavier distant est automatiquement désactivé.

---

# Dépannage

* **Aucun port COM visible** : vérifier le branchement de l’adaptateur, l’installation du pilote, puis cliquer sur **Refresh**.
* **Impossible d’ouvrir un port COM** : essayer un autre port, vérifier que le port n’est pas déjà utilisé par un autre logiciel.
* **Le X-07 ne réagit pas (BASIC/ASM/REMOTE KEYBOARD)** :
  * vérifier que le X-07 est bien en **SLAVE mode**,
  * vérifier RX / TX / GND,
  * ajuster **CHAR(s)** / **LINE(s)** pour la saisie BASIC.
* **LOAD"COM:" ne charge rien (CAS/K7)** :
  * vérifier que l’envoi est déclenché après le RETURN côté X-07,
  * vérifier la cohérence du fichier et l’offset d’envoi (0x0010).
* **SAVE"COM:" ne génère rien côté PC** :
  * vérifier que vous n’êtes pas en mode *slave*,
  * cliquer sur **Receive raw (SAVE "COM:")** avant de lancer `SAVE"COM:"`,
  * vérifier le câblage et la vitesse **Typing baud (8N2)**,
  * si le programme est très court, augmenter légèrement le timeout de capture côté PC (valeur interne).
* **Problèmes avec câble RTS/CTS** :
  * vérifier que **RTS et CTS sont correctement croisés**
  * vérifier que l’option **RTS/CTS cable** est activée
  * vérifier que l’adaptateur supporte le **hardware flow control**

---

# 🇬🇧 English

### Overview

This tool, based on a **Tkinter GUI**, allows transferring programs to a **Canon X-07** via a **serial connection**.

It supports:

* **BASIC (text listing)**: sending a `.txt` / `.bas` file by “typing” it line by line
* **BASIC (cassette raw stream)**: sending *and* receiving `.cas` / `.k7` files via `LOAD"COM:"` and `SAVE"COM:"`
* **ASM (fast)**: transferring a loader, then sending a `.bin` binary quickly
* **Remote keyboard (REMOTE KEYBOARD)**: typing from the PC directly to the X-07

The application provides:

* a timestamped log console `[hh:mm:ss]`
* a progress bar
* a cancel button
* special key buttons

---

### Requirements

* Python 3.10 or newer
* `pyserial` (automatically installed if needed)

---

### Required hardware

* A working Canon X-07
* USB ↔ serial adapter (FT232, CP2102, CH340, etc.)

Optional:

* cable with **RTS/CTS**

⚠️ Canon X-07 specific characteristic:
the **RX, TX, RTS and CTS signals are inverted**.

Example with FT232RL (FT_Prog):

```
Invert TXD : ON
Invert RXD : ON
Invert RTS : ON
Invert CTS : ON
```

⚠️ On some setups, add **pull-down resistors (~10 kΩ)** on:

* RX
* CTS

These resistors ensure a **stable logic level** when the lines are not actively driven by the Canon X-07.

---

### Running the application

```bash
python x07_loader.pyw
```

1. Select the **COM port** in *Serial settings*.
2. Adjust if needed:

* **Typing baud (8N2)**: speed used for BASIC typing, remote keyboard, and cassette streams
* **Loader baud (8N2)**: speed used for fast ASM transfer
* **CHAR(s)** / **LINE(s)**: delays (in seconds) between characters / lines during BASIC typing
* **RTS/CTS cable**: enables **hardware flow control RTS/CTS** if your cable supports it

---

### ⚠️ Canon X-07 “Slave mode” (required)

Must be enabled for:

* **BASIC text** (`.txt` / `.bas`)
* **ASM** (One click: fast loader + `.bin`)
* **REMOTE KEYBOARD**

In this mode:

* the serial device (`COM:`) becomes the **console input**
* the local keyboard is ignored
* received characters are interpreted as keyboard input

#### Entering slave mode

Type on the Canon X-07:

```basic
INIT#5,"COM:
EXEC&HEE1F
```

#### Leaving slave mode

Two methods:

1. Power cycle the Canon X-07
2. Send via serial:

```basic
EXEC&HEE33
```

The **Disable slave (EXEC&HEE33)** button performs this remotely.

---

# BASIC

### 1) BASIC text (.txt / .bas) via SLAVE mode

Goal: transfer a BASIC listing by “typing” it as if entered on the keyboard.

Procedure:

1. Put the X-07 in **SLAVE mode**
2. In **Text listing (.txt/.bas)**:

   * **Select .txt/.bas…**
   * **Send BASIC**

Operation:

* lines are sent in **8N2** with **CHAR(s)** and **LINE(s)** delays
* at the end, the tool sends `EXEC&HEE33` to release the remote console state

---

### 2) BASIC cassette stream (.cas / .k7) via LOAD "COM:" / SAVE "COM:" (raw bytes)

Goal: exchange a “cassette-style” program as a raw byte stream over serial, without slave mode.

---

#### a) Sending (PC → X-07) with `LOAD"COM:"`

On the Canon X-07:

1. **Do not** be in slave mode
2. Run:

```basic
LOAD"COM:"
```

3. Press **RETURN** (the X-07 waits for incoming bytes)

On the PC:

1. In **Cassette stream (.cas/.k7)**:

   * **Select .cas/.k7…**
   * (optional) **Inspect header**
   * **Send raw (LOAD "COM:")**

Important notes:

* transmission is done in **8N2**
* transfer starts at offset **`0x0010`** (validated base)
* **Inspect header** shows only a **preview @0x0000**

---

#### b) Receiving (X-07 → PC) with `SAVE"COM:"`

On the Canon X-07:

1. **Do not** be in slave mode
2. Run:

```basic
SAVE"COM:"
```

(or `SAVE"COM:name"`), then press **RETURN**

On the PC:

1. In **Cassette stream (.cas/.k7)**:

   * **Receive raw (SAVE "COM:")**
   * choose output `.cas` file

Operation:

* capture runs in **8N2**
* stops automatically after inactivity
* saved file contains:

  * a **16-byte header** (10×`D3` + 6-character name from filename)
  * followed by the **raw received stream**

---

# ASM (fast loader via COM)

Goal: transfer an ASM `.bin` quickly using an optimized loader.

Principle:

1. A loader is sent via `LOAD"COM:"`
2. The loader switches the serial link to high speed (e.g. 8000 baud)
3. The PC sends the binary as ASCII hex
4. The loader copies the data into memory and executes the program

---

## Methods

### Send ASM loader

Send only the loader.

On the X-07:

```basic
LOAD"COM:"
```

Then on PC:

* click **Send ASM loader**

The loader is transferred as a raw cassette stream.

---

### Send ASM (loader running)

Send only the binary once the loader is active.

* Uses **Loader baud (8N2)**
* Requires the loader to be running (`RUN`)
* No software delays (fast transfer)

⚠️ Important:

A synchronization prefix (“primer”) is automatically sent to improve reliability.

---

### One click: loader + ASM

Automatic sequence:

1. `LOAD"COM:"`
2. send loader
3. `EXEC&HEE33:RUN`
4. send binary

---

## Addresses

* **Loader addr**: loader memory address (e.g. `0x1800`)
* **ASM addr**: program memory address (e.g. `0x2000`)

⚠️ Must match:

```asm
ORG $2000
```

---

## Load address

### Loader addr

Defines where the loader is copied (default: `0x1800`)

### ASM addr

Defines where the `.bin` is copied (default: `0x2000`)

⚠️ Must match the assembly address.

Example:

```asm
ORG $2000
```

If mismatched:

* `JP` / `CALL`
* memory accesses
* referenced data

may break (crash or undefined behavior)

---

# Remote keyboard via SLAVE mode

Goal: control the X-07 like a terminal.

Requirements:

1. Enable **SLAVE mode**
2. Enable **REMOTE KEYBOARD: ON**

Operation:

* opens persistent **8N2** session
* sends ASCII characters directly
* buttons send special key codes

Safety:

* disabled automatically during transfers

---

# Troubleshooting

* **No COM port**: check adapter and drivers, click **Refresh**
* **Cannot open COM port**: try another port, ensure not in use
* **X-07 not responding**:
  * ensure SLAVE mode
  * check RX / TX / GND
  * adjust CHAR / LINE delays
* **LOAD"COM:" does nothing**:
  * start transfer after pressing RETURN
  * verify file and offset `0x0010`
* **SAVE"COM:" empty file**:
  * not in slave mode
  * click receive before SAVE
  * check wiring and baud
  * increase timeout if needed
* **RTS/CTS issues**:
  * verify wiring
  * enable option
  * ensure adapter supports it