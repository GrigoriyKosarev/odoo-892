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
        // console.log('[bio_excel] Import from Excel clicked');
        this.env.model.action.doAction({
            name: _t('Import from Excel'),
            type: 'ir.actions.act_window',
            res_model: 'bio.mrp.production.schedule.import.wizard',
            views: [[false, 'form']],
            target: 'new',
        }, {
            onClose: () => this.env.model.load(),
        });
    },

    async _onClickExportExcel(ev) {
        const orm = this.env.services.orm;
        try {
            // Get current MPS state context
            const context = {
                ...this.env.model.state.context,
            };

            // Call export method - this will create and download Excel file
            const action = await orm.call(
                'mrp.production.schedule',
                'action_export_product_demand',
                [[]],
                { context: context }
            );

            if (action && action.url) {
                window.location.href = action.url;
            } else if (action) {
                await this.env.services.action.doAction(action);
            }
        } catch (error) {
            this.env.services.notification.add(
                error.message || _t('Export failed'),
                { type: 'danger' }
            );
        }
    }
});

console.log('[bio_excel] MrpMpsControlPanel patched successfully');

