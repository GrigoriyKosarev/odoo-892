/** @odoo-module **/

import { registry } from "@web/core/registry";
import { _t } from "@web/core/l10n/translation";

/**
 * Service to inject "Import from Excel" button into MPS view
 * Uses same pattern as standard MPS code (_createProduct method)
 */
const MPSButtonService = {
    dependencies: ["action"],

    start(env, { action }) {
        console.log('[bio_excel] MPS Button Service started');

        let buttonInjected = false;

        // Function to inject button into control panel
        function injectButton() {
            // Check if we're on MPS client action
            const isMPSView = document.querySelector('.o_mrp_mps') ||
                              document.querySelector('[data-js-class="mrp_mps"]') ||
                              window.location.href.includes('mrp_mps');

            if (!isMPSView) {
                buttonInjected = false;
                return;
            }

            // Check if button already exists
            if (document.querySelector('.o_mps_import_excel')) {
                return;
            }

            // Find the "Add a Product" button first to insert after it
            const addProductBtn = Array.from(document.querySelectorAll('button'))
                .find(btn => btn.textContent.trim().includes('Add a Product') ||
                            btn.textContent.trim().includes('Додати продукт'));

            let controlPanel;
            if (addProductBtn) {
                // Insert right after "Add a Product" button
                controlPanel = addProductBtn.parentElement;
            } else {
                // Fallback: find control panel
                controlPanel = document.querySelector('.o_control_panel .o_cp_buttons') ||
                              document.querySelector('.o_control_panel .o_cp_action_menus') ||
                              document.querySelector('.o_control_panel');
            }

            if (!controlPanel) {
                console.log('[bio_excel] Control panel not found');
                return;
            }

            // Create Import button
            const importButton = document.createElement('button');
            importButton.className = 'btn btn-secondary o_mps_import_excel';
            importButton.type = 'button';
            importButton.innerHTML = '<i class="fa fa-download"></i> Import from Excel';

            // Use same pattern as _createProduct() from standard code
            importButton.addEventListener('click', () => {
                console.log('[bio_excel] Opening import wizard');
                action.doAction({
                    name: _t('Import from Excel'),
                    type: 'ir.actions.act_window',
                    res_model: 'bio.mrp.production.schedule.import.wizard',
                    views: [[false, 'form']],
                    target: 'new',
                });
            });

            // Insert button
            if (addProductBtn) {
                // Insert right after "Add a Product"
                addProductBtn.parentNode.insertBefore(importButton, addProductBtn.nextSibling);
                importButton.style.marginLeft = '4px';
            } else {
                // Append to control panel
                controlPanel.appendChild(importButton);
            }

            buttonInjected = true;
            console.log('[bio_excel] Import button injected successfully');
        }

        // Initial injection
        setTimeout(injectButton, 100);
        setTimeout(injectButton, 500);
        setTimeout(injectButton, 1000);

        // Watch for DOM changes (when switching views)
        const observer = new MutationObserver(() => {
            if (!buttonInjected) {
                injectButton();
            }
        });

        observer.observe(document.body, {
            childList: true,
            subtree: true
        });

        // Re-check periodically (fallback)
        setInterval(() => {
            if (!buttonInjected) {
                injectButton();
            }
        }, 2000);
    },
};

registry.category("services").add("bio_excel_mps_button", MPSButtonService);

