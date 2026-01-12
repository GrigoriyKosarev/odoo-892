/** @odoo-module **/

import { MasterProductionScheduleController } from '@mrp_mps/mrp_production_schedule/master_production_schedule_controller';
import { patch } from "@web/core/utils/patch";

patch(MasterProductionScheduleController.prototype, 'bio_excel.MasterProductionScheduleController', {

    /**
     * Handle Import from Excel button click
     */
    onImportFromExcel() {
        this.model._importFromExcel();
    }
});
