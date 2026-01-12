/** @odoo-module **/

import { MasterProductionScheduleController } from '@mrp_mps/mrp_production_schedule/master_production_schedule_controller';
import { patch } from "@web/core/utils/patch";
import { onMounted } from "@odoo/owl";

patch(MasterProductionScheduleController.prototype, 'bio_excel.MasterProductionScheduleController', {

    setup() {
        this._super(...arguments);

        onMounted(() => {
            this._addImportButton();
        });
    },

    /**
     * Add Import from Excel button to the control panel
     */
    _addImportButton() {
        const createProductButton = this.el.querySelector('button[name="onCreateProduct"]');
        if (createProductButton && !this.el.querySelector('.o_mps_import_excel')) {
            const importButton = document.createElement('button');
            importButton.className = 'btn btn-secondary o_mps_import_excel ms-2';
            importButton.type = 'button';
            importButton.innerHTML = '<i class="fa fa-download"></i> Import from Excel';
            importButton.addEventListener('click', () => this.onImportFromExcel());
            createProductButton.parentNode.insertBefore(importButton, createProductButton.nextSibling);
        }
    },

    /**
     * Handle Import from Excel button click
     */
    onImportFromExcel() {
        this.model._importFromExcel();
    }
});
