/** @odoo-module **/

import { patch } from "@web/core/utils/patch";
import { _t } from "@web/core/l10n/translation";
import { MrpMpsControlPanel } from "@mrp_mps/search/mrp_mps_control_panel";

console.log('[bio_excel] Patching MrpMpsControlPanel...');

// Patch the control panel to add Import from Excel handler
patch(MrpMpsControlPanel.prototype, 'bio_excel.MrpMpsControlPanel', {
    /**
     * Handle click on "Import from Excel" button
     * Similar to _onClickCreate in mrp_mps_control_panel.js
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

