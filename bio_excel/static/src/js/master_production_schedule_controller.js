/** @odoo-module **/

import { MasterProductionScheduleController } from '@mrp_mps/mrp_production_schedule/master_production_schedule_controller';
import { patch } from "@web/core/utils/patch";
import { onMounted, onPatched } from "@odoo/owl";

patch(MasterProductionScheduleController.prototype, 'bio_excel.MasterProductionScheduleController', {

    setup() {
        this._super(...arguments);

        onMounted(() => {
            console.log('[bio_excel] MPS Controller mounted');
            this._addImportButton();
        });

        onPatched(() => {
            this._addImportButton();
        });
    },

    /**
     * Add Import from Excel button to the control panel
     */
    _addImportButton() {
        // Try multiple selectors to find the button
        let createProductButton = null;

        // Try finding in document root
        createProductButton = document.querySelector('button[name="onCreateProduct"]');

        // Try finding in this.el
        if (!createProductButton && this.el) {
            createProductButton = this.el.querySelector('button[name="onCreateProduct"]');
        }

        console.log('[bio_excel] createProductButton:', createProductButton);
        console.log('[bio_excel] this.el:', this.el);

        if (createProductButton) {
            // Check if button already exists
            const existingButton = document.querySelector('.o_mps_import_excel');
            if (existingButton) {
                console.log('[bio_excel] Button already exists, skipping');
                return;
            }

            const importButton = document.createElement('button');
            importButton.className = 'btn btn-secondary o_mps_import_excel ms-2';
            importButton.type = 'button';
            importButton.innerHTML = '<i class="fa fa-download"></i> Import from Excel';
            importButton.addEventListener('click', () => this.onImportFromExcel());

            createProductButton.parentNode.insertBefore(importButton, createProductButton.nextSibling);
            console.log('[bio_excel] Import button added successfully');
        } else {
            console.log('[bio_excel] Create Product button not found');
        }
    },

    /**
     * Handle Import from Excel button click
     */
    onImportFromExcel() {
        this.model._importFromExcel();
    }
});
