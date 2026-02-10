# Mapala

Mapala est un outil simple de **mapping template ↔ source** avec **concaténation de colonnes**. Il ne gère pas les zones ni l’agrégation de lignes : tout est fait sur une seule zone (tout le template).

## Utilisation (dans l’app)

1. **Charger le template** : clique dans la **zone de prévisualisation** (en bas) et choisis un fichier (`.xlsx`, `.ods`, `.csv`).
2. **Charger la source** : clique dans la **liste des colonnes source** (à gauche) et choisis un fichier.
3. **Régler les lignes** :
   - **Champs cible ligne** = ligne qui contient les labels techniques du template.
   - **Labels** = lignes additionnelles (ex: `1,2`) si besoin.
   - **En‑tête source** = ligne d’en‑tête de la source.
4. **Mapper** : dans la **première ligne de la preview**, choisis la colonne source pour chaque champ cible.
5. **Concat** : choisis `Concat…` ou clique l’icône **Concat** pour un champ cible, puis configure la concaténation à droite (ordre, préfixes, séparateur, déduplication, ignorer vides).
6. **Exporter** : bouton **Exporter** en bas à droite. Formats : **xlsx / ods / csv**.

## Lancer l’app (développeurs)

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .
mapala
```

## Lancer via .command (macOS)

Double‑clique `mapala.command` : il crée automatiquement une `.venv` locale, installe les dépendances puis lance l’app.

## Build

Windows (.exe):
```bash
pip install -e .[build,formats]
python build_exe.py
```

macOS (.app / .dmg):
```bash
pip install -e .[build,formats]
python build_macos_app.py
./build_macos_dmg.sh
```

## Releases

Les releases GitHub incluent :
- **Windows** : `MapalaPortable.zip` (contient `Mapala.exe` + DLL)
- **macOS** : `Mapala.dmg` et `Mapala.app.zip`
