/** @odoo-module **/

import { MasterProductionScheduleModel } from '@mrp_mps/mrp_production_schedule/master_production_schedule_model';
import { _t } from "@web/core/l10n/translation";
import { patch } from "@web/core/utils/patch";

patch(MasterProductionScheduleModel.prototype, 'bio_excel.MasterProductionScheduleModel', {

    /**
     * Open import wizard for Excel file
     * @private
     */
    _importFromExcel() {
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
