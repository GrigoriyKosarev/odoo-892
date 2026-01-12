/** @odoo-module **/

import { patch } from "@web/core/utils/patch";
import { _t } from "@web/core/l10n/translation";

// Wait for modules to load, then patch the control panel
setTimeout(() => {
    // Get the MrpMpsControlPanel from loaded modules
    const modules = window.odoo?.loader?.modules;

    if (!modules) {
        console.error('[bio_excel] Odoo loader not found');
        return;
    }

    let MrpMpsControlPanel = null;

    // Find the MrpMpsControlPanel in loaded modules
    for (const [moduleName, module] of modules.entries()) {
        if (moduleName.includes('mrp_mps_control_panel') || moduleName.includes('mrp_mps/search/mrp_mps_control_panel')) {
            if (module.MrpMpsControlPanel) {
                MrpMpsControlPanel = module.MrpMpsControlPanel;
                console.log('[bio_excel] Found MrpMpsControlPanel in module:', moduleName);
                break;
            }
        }
    }

    if (!MrpMpsControlPanel) {
        console.error('[bio_excel] MrpMpsControlPanel not found in loaded modules');
        console.log('[bio_excel] Available modules:', Array.from(modules.keys()).filter(k => k.includes('mrp')));
        return;
    }

    // Patch the control panel to add Import from Excel handler
    patch(MrpMpsControlPanel.prototype, 'bio_excel.MrpMpsControlPanel', {
        /**
         * Handle click on "Import from Excel" button
         * @private
         */
        _onClickImportExcel(ev) {
            console.log('[bio_excel] Import from Excel clicked');
            this.env.model.action.doAction({
                name: _t('Import from Excel'),
                type: 'ir.actions.act_window',
                res_model: 'bio.mrp.production.schedule.import.wizard',
                views: [[false, 'form']],
                target: 'new',
            }, {
                onClose: () => this.env.model.load(),
            });
        }
    });

    console.log('[bio_excel] MrpMpsControlPanel patched successfully');
}, 2000);
