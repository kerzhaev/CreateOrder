# Архив временных тестовых копий — 09.08.2026

Сюда перенесены папки, которые автоматические проверки VBA создавали в корне проекта:

- `_tmp_acceptance_test`
- `_tmp_enrollment_bridge_test`
- `_tmp_enrollment_compact_ui_test`
- `_tmp_enrollment_fizo_reference_test`
- `_tmp_enrollment_medal_reference_test`
- `_tmp_enrollment_tariff_reference_test`
- `_tmp_monthly_compile`
- `_tmp_operator_ui_compile`
- `_tmp_personnel_action_wizard_test`
- `_tmp_personnel_events_test`
- `_tmp_personnel_ribbon_test`
- `_tmp_reference_compile`
- `_tmp_reference_linear_compile`
- `_tmp_zp12_validation_test`
- `old_form_modules` — пустая устаревшая корневая папка.

Это одноразовые копии книги и результаты проверок, а не рабочие исходники. При следующем запуске соответствующего теста папка `_tmp_*` будет создана заново автоматически.

Для восстановления старого состояния достаточно переместить нужную папку обратно в корень проекта. `CreateOrderBackups` в этот архив не переносился: это действующий каталог резервных копий до импорта VBA.
