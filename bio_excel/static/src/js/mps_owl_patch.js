/** @odoo-module **/

import { patch } from "@web/core/utils/patch";
import { _t } from "@web/core/l10n/translation";
import { registry } from "@web/core/registry";

console.log('[bio_excel] MPS OWL Patch loading...');

// Wait for all modules to load, then try to patch
setTimeout(() => {
    try {
        // Try to find MPS components in window.odoo
        const modules = window.odoo?.loader?.modules;

        if (!modules) {
            console.error('[bio_excel] Odoo loader modules not found');
            return;
        }

        console.log('[bio_excel] Available modules:', Array.from(modules.keys()));

        // Try to find MPS model and controller
        let ModelClass = null;
        let ControllerClass = null;

        for (const [moduleName, module] of modules.entries()) {
            if (moduleName.includes('master_production_schedule') || moduleName.includes('mrp_mps')) {
                console.log('[bio_excel] Found MPS module:', moduleName, module);

                if (module.MasterProductionScheduleModel) {
                    ModelClass = module.MasterProductionScheduleModel;
                    console.log('[bio_excel] Found MasterProductionScheduleModel');
                }

                if (module.MasterProductionScheduleController) {
                    ControllerClass = module.MasterProductionScheduleController;
                    console.log('[bio_excel] Found MasterProductionScheduleController');
                }
            }
        }

        // Patch Model if found
        if (ModelClass) {
            patch(ModelClass.prototype, 'bio_excel.MasterProductionScheduleModel', {
                _importFromExcel() {
                    console.log('[bio_excel] _importFromExcel called');
                    this.mutex.exec(() => {
                        this.env.services.action.doAction({
                            name: _t('Import from Excel'),
                            type: 'ir.actions.act_window',
                            res_model: 'bio.mrp.production.schedule.import.wizard',
                            views: [[false, 'form']],
                            target: 'new',
                        }, {
                            onClose: () => this.load(),
                        });
                    });
                }
            });
            console.log('[bio_excel] Model patched successfully');
        }

        // Patch Controller if found
        if (ControllerClass) {
            patch(ControllerClass.prototype, 'bio_excel.MasterProductionScheduleController', {
                onImportFromExcel() {
                    console.log('[bio_excel] onImportFromExcel clicked');
                    this.model._importFromExcel();
                }
            });
            console.log('[bio_excel] Controller patched successfully');
        }

        if (!ModelClass && !ControllerClass) {
            console.warn('[bio_excel] MPS components not found in loaded modules');
        }

    } catch (error) {
        console.error('[bio_excel] Error patching MPS components:', error);
    }
}, 3000); // Wait 3 seconds for all modules to load
