# simpleCopy

**Copier, ajouter et gérer le texte efficacement avec NVDA**

**auteur:** chai chaimee  
**url:** https://github.com/chaichaimee/simpleCopy

---

## Description

**simpleCopy** est une extension légère pour NVDA qui simplifie la copie de texte, l'extraction d'URL et la gestion de l'historique de parole.

Cet outil vous aide à capturer et organiser rapidement les informations sans interrompre votre flux de travail. Que vous copiiez du texte, récupériez des liens Web ou enregistriez du contenu parlé, simpleCopy fournit des raccourcis clavier intuitifs qui fonctionnent parfaitement avec NVDA.

---

## Raccourcis clavier

Toutes les commandes utilisent un système de tapotement multiple. Appuyez sur la combinaison de touches une, deux ou trois fois en succession rapide pour effectuer différentes actions.

### CTRL+Maj+A — Capture d'URL et de liens

- **Une pression:** Copie l'URL de la page Web actuelle.
- **Deux pressions:** Copie l'URL de destination du lien hypertexte focalisé.

### CTRL+Maj+V — Copier, ajouter et gérer le presse-papiers

- **Une pression:** Copie le texte sélectionné. Si du texte existe déjà dans le presse-papiers, la nouvelle sélection y est ajoutée.
- **Deux pressions:** Copie le texte à partir de la position actuelle du curseur de revue. Cela fonctionne avec toute sélection effectuée à l'aide du curseur de revue de NVDA, y compris les sélections multilignes et la sélection de document complet.
- **Trois pressions:** Efface tout le contenu du presse-papiers.

### F9 — Capture et gestion de la parole

- **Une pression:** Copie la sortie vocale la plus récente de NVDA.
- **Deux pressions:** Ajoute la sortie vocale la plus récente au contenu existant du presse-papiers.
- **Trois pressions:** Copie toute la sortie vocale accumulée depuis la première pression de F9.

### Maj+F9 — Navigation dans l'historique de parole

- **Une pression:** Navigue vers l'élément précédent de l'historique de parole.
- **Deux pressions:** Navigue vers l'élément suivant de l'historique de parole.
- **Trois pressions:** Ouvre le fichier journal complet de l'historique de parole.

---

## Fonctionnalités

Voici comment chaque fonctionnalité fonctionne en pratique:

### 1. Copier l'URL d'une page Web

Appuyez sur **CTRL+Maj+A une fois** lorsque vous naviguez sur un site Web. L'URL de la page actuelle est copiée dans votre presse-papiers. NVDA le confirme en lisant l'URL copiée.

### 2. Extraire l'URL d'un lien hypertexte

Focalisez un lien et appuyez sur **CTRL+Maj+A deux fois**. L'URL de destination est extraite et copiée sans ouvrir le lien.

### 3. Copier et ajouter du texte

Sélectionnez du texte et appuyez sur **CTRL+Maj+V une fois**. Si le presse-papiers est vide, le texte est copié. Si le presse-papiers contient déjà du texte, la nouvelle sélection est ajoutée avec un saut de ligne.

### 4. Copier depuis le curseur de revue

Utilisez le curseur de revue de NVDA pour sélectionner du texte (en utilisant NVDA+Maj+Flèche bas ou NVDA+CTRL+Maj+Flèche bas pour sélectionner plusieurs lignes), puis appuyez sur **CTRL+Maj+V deux fois**. Tout le texte sélectionné à partir de la position du curseur de revue est copié dans le presse-papiers. Cela fonctionne avec n'importe quelle taille de sélection, d'un seul mot à un document entier.

### 5. Effacer le presse-papiers

Appuyez sur **CTRL+Maj+V trois fois** pour effacer instantanément tout le contenu du presse-papiers. NVDA le confirme avec le message "Clean".

### 6. Copier la dernière parole

Lorsque NVDA dit quelque chose que vous souhaitez enregistrer, appuyez sur **F9 une fois**. La dernière phrase prononcée est copiée dans votre presse-papiers.

### 7. Ajouter la parole

Appuyez sur **F9 deux fois** pour ajouter la dernière phrase prononcée au contenu existant du presse-papiers.

### 8. Enregistrer l'historique de parole

Appuyez sur **F9 trois fois** pour copier toute la sortie vocale accumulée pendant votre session en cours.

### 9. Naviguer dans l'historique de parole

Utilisez **Maj+F9 une fois** pour reculer dans l'historique de parole, et **Maj+F9 deux fois** pour avancer. Cela vous permet de consulter les sorties vocales passées sans changer votre focus actuel.

### 10. Accéder au fichier journal de parole

Appuyez sur **Maj+F9 trois fois** pour ouvrir le fichier complet de l'historique de parole dans votre éditeur de texte par défaut pour le consulter, le rechercher ou le copier.

### 11. Intelligence contextuelle

Lorsque vous tapez dans des champs modifiables, simpleCopy n'interfère pas. Les commandes ne s'activent que lorsqu'elles sont utiles, préservant ainsi votre flux de travail normal.

---

## Soutenez-moi

Si cette extension vous aide à travailler plus efficacement, envisagez de faire un petit don pour soutenir le développement futur.

[![Soutenez-moi](https://img.shields.io/badge/Donate-Support%20Me-blue?style=for-the-badge&logo=stripe)](https://buy.stripe.com/dRm9AU1xQ3Ds22N6VK1VK01)

Votre soutien aide à garder ce projet vivant et en amélioration.

---

© 2026 Chai Chaimee Extension NVDA Publiée sous licence GNU General Public License