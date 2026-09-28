# Catalogue des outils LLM

Catalogue vérifié le 28 septembre 2026 contre le lot IDE courant : **180 outils**. Cet inventaire est extrait de `LlmVbeTools.Definitions`, y compris les catalogues Git et Editor. Il décrit les outils exposés au modèle, et ne prétend pas inventorier toutes les commandes du pont.

## Contrat

- Les noms et paramètres JSON sont sensibles à la casse.
- `Project` désigne un projet résolu dans le VBE vivant ; lire son identité avant une action.
- SHA, versions de projet, arbre, volet ou presse-papiers sont relus avant les mutations correspondantes.
- « Inspection » signifie autorisé en Discussion/Plan selon `ReadOnlyTools` ; certaines inspections ouvrent une fenêtre ou compilent le projet et ont donc un effet local.
- « Action » nécessite le mode Agent et reste soumise à la politique configurée et aux gardes de la commande.
- Git et les opérations natives asynchrones passent par `InvokeAsync`. Une erreur ou un retour « pending » ne justifie pas de répéter une mutation.

Les descriptions, les champs facultatifs, les types et les bornes font autorité dans [le catalogue principal](../../src/CodexVBE/Llm/Chat/LlmVbeTools.cs), [Editor](../../src/CodexVBE/Llm/Chat/LlmVbeTools.Editor.cs) et [Git](../../src/CodexVBE/Llm/Chat/LlmVbeTools.Git.cs).

## Outils exposés

| Outil | Catégorie de permission | Paramètres requis |
| --- | --- | --- |
| `add_form_control` | Action | `Project`, `Form`, `ExpectedFormVersion`, `ControlType`, `Control`, `Left`, `Top`, `Width`, `Height` |
| `add_form_page` | Action | `Project`, `Form`, `ParentPath`, `NewName`, `ExpectedTreeVersion` |
| `add_form_tab` | Action | `Project`, `Form`, `ParentPath`, `NewName`, `ExpectedTreeVersion` |
| `add_nested_form_control` | Action | `Project`, `Form`, `ParentPath`, `ExpectedTreeVersion`, `ControlType`, `Control`, `Left`, `Top`, `Width`, `Height` |
| `add_reference_file` | Action | `Project`, `ExpectedReferencesVersion`, `Path` |
| `add_reference_guid` | Action | `Project`, `ExpectedReferencesVersion`, `Guid`, `Major`, `Minor` |
| `add_toolbar_command` | Action | `ObjectName`, `ExpectedToolbarControlsVersion`, `ControlId`, `ControlCaption` |
| `add_watch` | Action | `Project`, `Module`, `ExpectedMode`, `Expression` |
| `apply_code_edit` | Action | `Project`, `Module`, `ExpectedSha256`, `StartLine`, `Count`, `Action` |
| `apply_form_layout` | Action | `Project`, `Form`, `ExpectedTreeVersion`, `Items`, `Action` |
| `apply_local_rename` | Action | `Project`, `Module`, `Procedure`, `ProcKind`, `ExpectedSha256`, `ExpectedMode`, `StartLine`, `StartColumn`, `Query`, `NewName` |
| `apply_parameter_rename` | Action | `Project`, `Module`, `Procedure`, `ProcKind`, `ExpectedSha256`, `ExpectedMode`, `StartLine`, `StartColumn`, `Query`, `NewName` |
| `arrange_editor_windows` | Action | `Action`, `ExpectedWindowVersion`, `ControlCaption` |
| `certificate_trust` | Inspection | `CertificateThumbprint` |
| `close_vbe_window` | Action | `WindowCaption`, `WindowType` |
| `code_bookmark` | Inspection | `Project`, `Action` |
| `code_pane_layout` | Inspection | `Project`, `Module` |
| `code_panes` | Inspection | — |
| `compile_project` | Inspection | `Project`, `ExpectedMode` |
| `component_properties` | Inspection | `Project`, `Module` |
| `component_property_value` | Inspection | `Project`, `Module`, `Property` |
| `copy_code` | Action | `Project`, `Module`, `ExpectedSha256`, `StartLine`, `StartColumn`, `EndLine`, `EndColumn` |
| `create_class` | Action | `Project`, `Module`, `ExpectedMode` |
| `create_event_procedure` | Action | `Project`, `Form`, `ObjectName`, `EventName`, `ExpectedSha256`, `ExpectedTreeVersion` |
| `create_form` | Action | `Project`, `Form` |
| `create_module` | Action | `Project`, `Module`, `ExpectedMode` |
| `create_procedure` | Action | `Project`, `Module`, `Procedure`, `ProcKind`, `Text`, `ExpectedSha256` |
| `create_toolbar` | Action | `ObjectName`, `ExpectedToolbarCollectionVersion` |
| `cut_code` | Action | `Project`, `Module`, `ExpectedSha256`, `StartLine`, `StartColumn`, `EndLine`, `EndColumn` |
| `debug_dialog` | Inspection | — |
| `debug_global` | Action | `Project`, `ExpectedMode`, `Action` |
| `debug_item` | Inspection | `Pane`, `Action`, `PathSegments` |
| `debug_state` | Inspection | `Project` |
| `debug_windows` | Inspection | — |
| `edit_watch` | Action | `Project`, `ExpectedMode`, `Expression`, `Context`, `NewExpression` |
| `editor_layout` | Inspection | — |
| `export_component` | Action | `Project`, `Module`, `ExpectedComponentVersion`, `Path` |
| `find_code` | Inspection | `Project`, `Query` |
| `focus_vbe_window` | Inspection | `WindowCaption`, `WindowType` |
| `form_clipboard_state` | Inspection | `Project`, `Form` |
| `form_control_properties` | Inspection | `Project`, `Form`, `Control` |
| `form_event_catalog` | Inspection | `Project`, `Form` |
| `form_list_items` | Inspection | `Project`, `Form`, `ControlPath` |
| `form_properties` | Inspection | `Project`, `Form` |
| `form_run_status` | Inspection | `Project`, `Query` |
| `form_state` | Inspection | `Project`, `Form` |
| `form_tree` | Inspection | `Project`, `Form` |
| `git_branch_create` | Action | `Project`, `ExpectedState`, `Name` |
| `git_branch_switch` | Action | `Project`, `ExpectedState`, `Name` |
| `git_branch_track` | Action | `Project`, `ExpectedState`, `Name` |
| `git_branches` | Inspection | `Project` |
| `git_checkpoint_create` | Action | `Project`, `ExpectedState`, `Name` |
| `git_checkpoint_restore` | Action | `Project`, `ExpectedState`, `Name` |
| `git_checkpoints` | Inspection | `Project` |
| `git_commit` | Action | `Project`, `ExpectedState`, `Text` |
| `git_commit_read` | Inspection | `Project`, `Name` |
| `git_commit_selected` | Action | `Project`, `ExpectedState`, `Name`, `Text`, `Choice` |
| `git_conflict_read` | Inspection | `Project`, `Path` |
| `git_conflicts` | Inspection | `Project` |
| `git_fetch` | Action | `Project`, `ExpectedState` |
| `git_history` | Inspection | `Project` |
| `git_merge_abort` | Action | `Project`, `ExpectedState` |
| `git_merge_begin` | Action | `Project`, `ExpectedState`, `Name` |
| `git_merge_complete` | Action | `Project`, `ExpectedState`, `Text` |
| `git_merge_resolve` | Action | `Project`, `ExpectedState`, `Path`, `Choice`, `Text` |
| `git_module_restore` | Action | `Project`, `ExpectedState`, `Name`, `Path` |
| `git_pr_prepare` | Action | `Project`, `ExpectedState`, `Name`, `Text`, `Choice` |
| `git_pull` | Action | `Project`, `ExpectedState` |
| `git_pull_requests` | Inspection | `Project` |
| `git_push` | Action | `Project`, `ExpectedState` |
| `git_remote_branches` | Action | `Project`, `ExpectedState` |
| `git_rollback` | Action | `Project`, `ExpectedState` |
| `git_status` | Inspection | `Project` |
| `immediate_execute` | Action | `Project`, `ExpectedMode`, `Text` |
| `import_component` | Action | `Project`, `ExpectedProjectVersion`, `Path` |
| `insert_code_file` | Action | `Project`, `Module`, `Path`, `StartLine`, `ExpectedSha256` |
| `inspect_code_file` | Inspection | `Path` |
| `invoke_debug` | Action | `Project`, `Module`, `ExpectedSha256`, `StartLine`, `ExpectedMode`, `Action`, `ControlId`, `ControlCaption` |
| `link_vbe_window` | Action | `WindowCaption`, `WindowType`, `ExpectedWindowVersion`, `Action` |
| `list_addins` | Inspection | — |
| `list_commands` | Inspection | — |
| `list_form_control_types` | Inspection | — |
| `list_forms` | Inspection | `Project` |
| `list_modules` | Inspection | `Project` |
| `list_object_browser` | Inspection | `Pane` |
| `list_procedures` | Inspection | `Project`, `Module` |
| `list_projects` | Inspection | — |
| `list_reference_types` | Inspection | `Project`, `Guid`, `Major`, `Minor` |
| `list_references` | Inspection | `Project` |
| `list_signing_certificates` | Inspection | — |
| `list_toolbars` | Inspection | — |
| `list_type_members` | Inspection | `Project`, `Guid`, `Major`, `Minor`, `TypeIndex`, `TypeIdentity` |
| `native_code_history` | Action | `Project`, `Action`, `ExpectedMode`, `ExpectedProjectVersion`, `ControlCaption` |
| `native_code_history_state` | Inspection | `Project` |
| `native_code_navigation` | Action | `Action` |
| `native_form_clipboard` | Action | `Project`, `Form`, `Action`, `ExpectedDesignerSelectionVersion`, `ExpectedClipboardVersion` |
| `native_form_history` | Action | `Project`, `Form`, `Action`, `ExpectedTreeVersion` |
| `navigate_code` | Inspection | `Project`, `Action` |
| `open_debug_pane` | Inspection | `Action` |
| `open_form` | Inspection | `Project`, `Form` |
| `open_object_browser` | Inspection | — |
| `paste_code` | Action | `Project`, `Module`, `ExpectedSha256`, `ExpectedClipboardVersion`, `StartLine`, `StartColumn`, `EndLine`, `EndColumn` |
| `preview_code_edit` | Inspection | `Project`, `Module`, `ExpectedSha256`, `StartLine`, `Count`, `Action` |
| `preview_form_layout` | Inspection | `Project`, `Form`, `ExpectedTreeVersion`, `Items`, `Action` |
| `preview_local_rename` | Inspection | `Project`, `Module`, `Procedure`, `ProcKind`, `ExpectedSha256`, `StartLine`, `StartColumn`, `Query`, `NewName` |
| `preview_parameter_rename` | Inspection | `Project`, `Module`, `Procedure`, `ProcKind`, `ExpectedSha256`, `StartLine`, `StartColumn`, `Query`, `NewName` |
| `procedure_run_status` | Inspection | `Project`, `Query` |
| `project_persistence_status` | Inspection | `Project` |
| `project_properties` | Inspection | `Project` |
| `project_signature_status` | Inspection | `Project` |
| `project_symbols` | Inspection | `Project` |
| `quick_watch` | Action | `Project`, `Module`, `ExpectedSha256`, `ExpectedMode`, `StartLine`, `StartColumn`, `EndColumn`, `Expression` |
| `read_code_clipboard` | Inspection | — |
| `read_debug_options` | Inspection | — |
| `read_module` | Inspection | `Project`, `Module` |
| `read_object_browser` | Inspection | — |
| `read_project_signature_dialog` | Inspection | `Project`, `ExpectedMode` |
| `read_runtime_forms` | Inspection | — |
| `read_user_file` | Inspection | `Path` |
| `read_vbe_options` | Inspection | — |
| `recover_form_cut` | Action | `Project`, `Form`, `DesignerClipboardRecoveryId`, `ExpectedDesignerSelectionVersion`, `ExpectedClipboardVersion` |
| `redo_code_edit` | Action | `Project`, `Module`, `ExpectedSha256` |
| `remove_component` | Action | `Project`, `Module`, `ExpectedProjectVersion`, `ExpectedComponentVersion` |
| `remove_form_control` | Action | `Project`, `Form`, `ControlPath`, `ExpectedTreeVersion` |
| `remove_form_page_tab` | Action | `Project`, `Form`, `ControlPath`, `ExpectedTreeVersion` |
| `remove_procedure` | Action | `Project`, `Module`, `Procedure`, `ProcKind`, `ExpectedSha256` |
| `remove_reference` | Action | `Project`, `ExpectedReferencesVersion`, `Guid`, `Major`, `Minor` |
| `remove_toolbar` | Action | `ObjectName`, `ExpectedToolbarCollectionVersion`, `ExpectedToolbarControlsVersion` |
| `remove_toolbar_command` | Action | `ObjectName`, `ExpectedToolbarControlsVersion`, `ControlId`, `ControlCaption`, `InsertIndex` |
| `remove_watch` | Action | `Project`, `ExpectedMode`, `Expression`, `Context` |
| `rename_component` | Action | `Project`, `Module`, `ExpectedComponentVersion`, `NewName` |
| `rename_form_control` | Action | `Project`, `Form`, `ExpectedFormVersion`, `Control`, `NewName` |
| `replace_lines` | Action | `Project`, `Module`, `ExpectedSha256`, `StartLine`, `Count`, `Text` |
| `replace_procedure` | Action | `Project`, `Module`, `Procedure`, `ProcKind`, `Text`, `ExpectedSha256` |
| `respond_debug_dialog` | Action | `Diagnostic`, `Button` |
| `restore_form_clipboard` | Action | `Project`, `Form`, `DesignerClipboardRecoveryId`, `ExpectedDesignerSelectionVersion`, `ExpectedClipboardVersion` |
| `run_form` | Action | `Project`, `Form`, `ExpectedMode`, `ExpectedSha256`, `ExpectedTreeVersion`, `ControlCaption` |
| `run_procedure` | Action | `Project`, `Module`, `Procedure`, `ExpectedSha256`, `ExpectedMode`, `Arguments` |
| `run_sub` | Action | `Project`, `Module`, `Procedure`, `ExpectedSha256`, `ExpectedMode` |
| `save_host_document` | Action | `Project`, `ExpectedProjectVersion`, `ExpectedHostPath` |
| `save_host_document_as` | Action | `Project`, `ExpectedProjectVersion`, `Path` |
| `scroll_code_pane` | Action | `Project`, `Module`, `Pane`, `ExpectedWindowVersion`, `ExpectedSha256`, `ExpectedMode`, `StartLine` |
| `select_code` | Inspection | `Project`, `Module`, `ExpectedSha256`, `StartLine` |
| `select_code_range` | Inspection | `Project`, `Module`, `ExpectedSha256`, `StartLine`, `StartColumn`, `EndLine`, `EndColumn` |
| `select_form_controls` | Action | `Project`, `Form`, `Items`, `ExpectedDesignerSelectionVersion` |
| `select_object_browser` | Inspection | — |
| `select_procedure` | Inspection | `Project`, `Module`, `Procedure`, `ProcKind`, `ExpectedSha256` |
| `set_addin_connection` | Action | `ProgId`, `ExpectedAddInVersion`, `Action` |
| `set_class_instancing` | Action | `Project`, `Module`, `ExpectedComponentVersion`, `Value` |
| `set_code_split` | Action | `Project`, `Module`, `ExpectedSha256`, `StartLine`, `ExpectedMode`, `Action`, `ControlCaption` |
| `set_code_view` | Action | `Project`, `Module`, `Pane`, `ExpectedWindowVersion`, `ExpectedSha256`, `ExpectedMode`, `StartLine`, `Action` |
| `set_component_property` | Action | `Project`, `Module`, `ExpectedComponentVersion`, `Property`, `Value` |
| `set_form_control_caption` | Action | `Project`, `Form`, `ExpectedFormVersion`, `Control`, `Caption` |
| `set_form_control_font` | Action | `Project`, `Form`, `ExpectedFormVersion`, `Control`, `FontName`, `FontSize`, `FontBold` |
| `set_form_control_geometry` | Action | `Project`, `Form`, `ExpectedFormVersion`, `Control`, `Left`, `Top`, `Width`, `Height` |
| `set_form_list_binding` | Action | `Project`, `Form`, `ControlPath`, `ExpectedHostPath`, `ExpectedTreeVersion`, `ExpectedSha256`, `SheetName`, `RangeAddress` |
| `set_form_list_initializer` | Action | `Project`, `Form`, `ControlPath`, `ExpectedTreeVersion`, `ExpectedSha256` |
| `set_form_node_picture` | Action | `Project`, `Form`, `ControlPath`, `ExpectedTreeVersion`, `Property`, `Path` |
| `set_form_node_property` | Action | `Project`, `Form`, `ControlPath`, `ExpectedTreeVersion`, `Property`, `Value` |
| `set_form_picture` | Action | `Project`, `Form`, `ExpectedFormVersion`, `Path` |
| `set_form_property` | Action | `Project`, `Form`, `ExpectedFormVersion`, `Property`, `Value` |
| `set_form_tab_order` | Action | `Project`, `Form`, `ExpectedTreeVersion`, `Items` |
| `set_project_property` | Action | `Project`, `ExpectedProjectVersion`, `Property`, `Value` |
| `set_toolbar_placement` | Action | `ObjectName`, `ExpectedToolbarLayoutVersion`, `Action` |
| `set_toolbar_position` | Action | `ObjectName`, `ExpectedToolbarLayoutVersion`, `Action` |
| `set_toolbar_visibility` | Action | `ObjectName`, `ExpectedWindowVersion`, `Action` |
| `set_vbe_option` | Action | `Pane`, `Property`, `Value`, `ExpectedOptionsVersion` |
| `set_window_bounds` | Action | `WindowCaption`, `WindowType`, `ExpectedWindowVersion`, `Left`, `Top`, `Width`, `Height` |
| `set_window_state` | Action | `WindowCaption`, `WindowType`, `ExpectedWindowVersion`, `Action` |
| `show_vbe_window` | Action | `WindowCaption`, `WindowType` |
| `sign_project` | Action | `Project`, `ExpectedProjectVersion`, `ExpectedMode`, `CertificateThumbprint` |
| `status` | Inspection | — |
| `toolbar_controls` | Inspection | `ObjectName` |
| `undo_code_edit` | Action | `Project`, `Module`, `ExpectedSha256` |
| `vbe_environment` | Inspection | — |
| `vbe_windows` | Inspection | — |
| `verify_vba_signature_file` | Inspection | `Path` |
| `window_layout` | Inspection | `WindowCaption`, `WindowType` |
| `window_linkage` | Inspection | `WindowCaption`, `WindowType` |
| `z_order_form_control` | Action | `Project`, `Form`, `ControlPath`, `ExpectedTreeVersion`, `ZPosition` |

## Limites et vérification

Consulter [l’état du projet](../project.md), [les travaux restants](../roadmap.md), [le concepteur](designer.md) et [la couverture mesurée](../test-coverage.md). Un outil exposé n’est pas une preuve de qualification dans tous les hôtes.

## Extensions fonctionnelles

Voir [les contrats, preuves et limites des nouvelles fonctions](../reference/functional-extensions.md). La personnalisation des barres accepte `Temporary=false` pour demander la persistance native ; celle-ci doit être relue après redémarrage.
