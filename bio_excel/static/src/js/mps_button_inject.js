/** @odoo-module **/

import { registry } from "@web/core/registry";

const MPSButtonService = {
    dependencies: ["action"],
    start(env, { action }) {
        console.log('[bio_excel] MPS Button Service started');

        // Function to inject button
        function injectButton() {
            // Check if button already exists
            if (document.querySelector('.o_mps_import_excel')) {
                return;
            }

            // Check if this is MPS view
            const isMPSView = document.querySelector('.o_mrp_mps') ||
                              window.location.href.includes('mrp_mps') ||
                              window.location.href.includes('action=980') ||
                              document.querySelector('[data-menu-xmlid="mrp_mps.menu_mrp_mps"]');

            if (!isMPSView) {
                return;
            }

            // Find control panel
            const controlPanel = document.querySelector('.o_control_panel .o_cp_buttons') ||
                                 document.querySelector('.o_control_panel') ||
                                 document.querySelector('.o_cp_top_right');

            if (!controlPanel) {
                console.log('[bio_excel] Control panel not found');
                return;
            }

            // Create button
            const importButton = document.createElement('button');
            importButton.className = 'btn btn-secondary o_mps_import_excel ms-2';
            importButton.type = 'button';
            importButton.innerHTML = '<i class="fa fa-download"></i> Import from Excel';

            importButton.addEventListener('click', () => {
                console.log('[bio_excel] Import button clicked');
                action.doAction({
                    name: 'Import from Excel',
                    type: 'ir.actions.act_window',
                    res_model: 'bio.mrp.production.schedule.import.wizard',
                    views: [[false, 'form']],
                    target: 'new',
                });
            });

            controlPanel.appendChild(importButton);
            console.log('[bio_excel] Import button injected successfully');
        }

        // Try to inject button on load and on URL changes
        injectButton();

        // Watch for URL changes
        let lastUrl = location.href;
        new MutationObserver(() => {
            const url = location.href;
            if (url !== lastUrl) {
                lastUrl = url;
                setTimeout(injectButton, 500);
            }
        }).observe(document, { subtree: true, childList: true });

        // Also try periodically
        setInterval(injectButton, 2000);
    },
};

registry.category("services").add("bio_excel_mps_button", MPSButtonService);
