import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class StandaloneTooltipComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _tooltipElement: HTMLDivElement | null = null;
    private _iconElement: HTMLElement | null = null;
    
    // State management
    private _hideTimeout: number | undefined = undefined;
    private _showTimeout: number | undefined = undefined;
    private _isTooltipVisible = false;
    private _isHovering = false;
    private _componentId: string;
    private _positioningRetries = 0;
    private _maxRetries = 6;
    private _positioned = false;
    private _observerTimeout: number | undefined = undefined;
    private _isHidden = false;

    constructor() {
        this._componentId = 'tooltip-' + Math.random().toString(36).substring(2, 9);
    }

    public init(
        context: ComponentFramework.Context<IInputs>, 
        notifyOutputChanged: () => void, 
        state: ComponentFramework.Dictionary, 
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;

        try {
            console.log("Initializing tooltip component:", this._componentId);
            
            // Check if component should be hidden
            this._isHidden = this.getBooleanParameter('isHidden', false);
            
            // ALWAYS hide the PCF container to eliminate visible space
            this.hideContainerCompletely();
            
            if (this._isHidden) {
                console.log("Component is set to hidden - not creating tooltip");
                return;
            }
            
            this.injectStyles();
            this.createTooltipElement();
            
            // Debug current state
            this.debugCurrentState();
            
            // Start positioning strategy
            this.startPositioningStrategy();
            
            console.log("Tooltip component initialized successfully");
        } catch (error) {
            console.error("Error during initialization:", error);
            // Only show fallback if not hidden
            if (!this._isHidden) {
                this.createFallbackIcon();
            }
        }
    }

    private hideContainerCompletely(): void {
        // Make the PCF container invisible but keep it in the DOM for field binding
        if (this._container) {
            this._container.style.cssText = `
                visibility: hidden !important;
                opacity: 0 !important;
                width: 1px !important;
                height: 1px !important;
                margin: 0 !important;
                padding: 0 !important;
                border: none !important;
                position: absolute !important;
                overflow: hidden !important;
                z-index: -1 !important;
                pointer-events: none !important;
                background: transparent !important;
            `;
            this._container.innerHTML = '';
            this._container.className = 'pcf-tooltip-hidden';
        }
    }

    private injectStyles(): void {
        const styleId = 'standalone-tooltip-styles-v7';
        if (document.getElementById(styleId)) return;

        const bgColor = this.getStringParameter('tooltipBackgroundColor', '#2d3748');
        const textColor = this.getStringParameter('tooltipTextColor', '#ffffff');
        const maxWidth = this.getNumberParameter('maxWidth', 350);

        const style = document.createElement('style');
        style.id = styleId;
        style.textContent = `
            /* Enhanced tooltip icon styles with better colors */
            .tooltip-icon-base {
                display: inline-flex !important;
                align-items: center !important;
                justify-content: center !important;
                width: 16px !important;
                height: 16px !important;
                cursor: help !important;
                border-radius: 50% !important;
                background: linear-gradient(135deg, #4299e1 0%, #3182ce 100%) !important;
                border: 1px solid #2b77cb !important;
                font-size: 10px !important;
                font-weight: 700 !important;
                color: #ffffff !important;
                vertical-align: middle !important;
                flex-shrink: 0 !important;
                transition: all 0.2s ease !important;
                z-index: 1000 !important;
                position: relative !important;
                box-shadow: 0 2px 4px rgba(66, 153, 225, 0.3) !important;
                font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, sans-serif !important;
            }
            
            .tooltip-icon-base:hover {
                background: linear-gradient(135deg, #3182ce 0%, #2c5282 100%) !important;
                transform: scale(1.1) !important;
                box-shadow: 0 4px 12px rgba(66, 153, 225, 0.4) !important;
                border-color: #2a69ac !important;
            }
            
            .tooltip-icon-base:focus {
                outline: 2px solid #4299e1 !important;
                outline-offset: 2px !important;
                background: linear-gradient(135deg, #3182ce 0%, #2c5282 100%) !important;
            }

            .tooltip-icon-base:active {
                transform: scale(0.95) !important;
                box-shadow: 0 1px 2px rgba(66, 153, 225, 0.2) !important;
            }

            /* Positioning styles */
            .tooltip-icon-inline {
                margin: 0 0 0 8px !important;
            }

            .tooltip-icon-absolute {
                position: absolute !important;
                right: 8px !important;
                top: 50% !important;
                transform: translateY(-50%) !important;
                margin: 0 !important;
                z-index: 100 !important;
            }

            .tooltip-icon-field-inline {
                margin: 0 0 0 8px !important;
                position: relative !important;
                top: 0 !important;
            }

            .tooltip-icon-label {
                margin: 0 0 0 6px !important;
                flex-shrink: 0 !important;
                vertical-align: baseline !important;
            }

            .tooltip-icon-field-absolute {
                position: absolute !important;
                right: 8px !important;
                top: 50% !important;
                transform: translateY(-50%) !important;
                z-index: 100 !important;
                margin: 0 !important;
            }

            .tooltip-icon-choice-absolute {
                position: absolute !important;
                right: -24px !important;
                top: 2px !important;
                z-index: 100 !important;
                margin: 0 !important;
            }

            /* Enhanced tooltip content with better styling */
            .standalone-tooltip-content {
                position: fixed !important;
                background: ${bgColor} !important;
                color: ${textColor} !important;
                padding: 16px 20px !important;
                border-radius: 12px !important;
                font-size: 14px !important;
                line-height: 1.5 !important;
                max-width: ${maxWidth}px !important;
                min-width: 200px !important;
                z-index: 999999 !important;
                box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04) !important;
                border: 1px solid rgba(255, 255, 255, 0.1) !important;
                visibility: hidden !important;
                opacity: 0 !important;
                transform: translateY(-8px) scale(0.95) !important;
                transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
                pointer-events: none !important;
                word-wrap: break-word !important;
                font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, sans-serif !important;
                box-sizing: border-box !important;
                backdrop-filter: blur(16px) !important;
                contain: layout style paint !important;
            }
            
            .standalone-tooltip-content.visible {
                visibility: visible !important;
                opacity: 1 !important;
                transform: translateY(0) scale(1) !important;
                pointer-events: auto !important;
            }

            /* Enhanced HTML content styling with image support */
            .standalone-tooltip-content p {
                margin: 0 0 8px 0 !important;
            }

            .standalone-tooltip-content p:last-child {
                margin-bottom: 0 !important;
            }

            .standalone-tooltip-content strong,
            .standalone-tooltip-content b {
                font-weight: 600 !important;
                color: inherit !important;
            }

            .standalone-tooltip-content em,
            .standalone-tooltip-content i {
                font-style: italic !important;
            }

            .standalone-tooltip-content ul,
            .standalone-tooltip-content ol {
                margin: 8px 0 !important;
                padding-left: 20px !important;
            }

            .standalone-tooltip-content li {
                margin: 4px 0 !important;
            }

            .standalone-tooltip-content br {
                line-height: 1.5 !important;
            }

            .standalone-tooltip-content img {
                max-width: 100% !important;
                height: auto !important;
                display: inline-block !important;
                vertical-align: middle !important;
                margin: 0 4px !important;
                border-radius: 4px !important;
            }

            .standalone-tooltip-content a {
                color: #60a5fa !important;
                text-decoration: underline !important;
            }

            .standalone-tooltip-content a:hover {
                color: #93c5fd !important;
            }

            .standalone-tooltip-content code {
                background: rgba(255, 255, 255, 0.1) !important;
                padding: 2px 4px !important;
                border-radius: 3px !important;
                font-family: 'Courier New', monospace !important;
                font-size: 0.9em !important;
            }

            /* Tooltip arrows with enhanced styling */
            .standalone-tooltip-content::before {
                content: '' !important;
                position: absolute !important;
                border: 8px solid transparent !important;
                filter: drop-shadow(0 -2px 4px rgba(0, 0, 0, 0.1)) !important;
            }

            .standalone-tooltip-content.bottom::before {
                top: -16px !important;
                left: 50% !important;
                transform: translateX(-50%) !important;
                border-bottom-color: ${bgColor} !important;
            }

            .standalone-tooltip-content.top::before {
                bottom: -16px !important;
                left: 50% !important;
                transform: translateX(-50%) !important;
                border-top-color: ${bgColor} !important;
            }

            .standalone-tooltip-content.left::before {
                right: -16px !important;
                top: 50% !important;
                transform: translateY(-50%) !important;
                border-left-color: ${bgColor} !important;
            }

            .standalone-tooltip-content.right::before {
                left: -16px !important;
                top: 50% !important;
                transform: translateY(-50%) !important;
                border-right-color: ${bgColor} !important;
            }

            /* Ensure PCF containers are minimally visible but take no space */
            .pcf-tooltip-container,
            .pcf-tooltip-hidden {
                visibility: hidden !important;
                opacity: 0 !important;
                width: 1px !important;
                height: 1px !important;
                margin: 0 !important;
                padding: 0 !important;
                overflow: hidden !important;
                position: absolute !important;
                z-index: -1 !important;
                pointer-events: none !important;
                background: transparent !important;
            }

            /* Power Platform specific adjustments */
            .ms-TextField-wrapper.has-tooltip,
            .ms-Dropdown-container.has-tooltip {
                display: inline-flex !important;
                align-items: center !important;
                flex: 1 !important;
                min-width: 0 !important;
            }

            .ms-TextField .tooltip-icon-absolute {
                right: 12px !important;
                z-index: 10 !important;
            }

            .ms-Dropdown .tooltip-icon-absolute {
                right: 32px !important;
                z-index: 10 !important;
            }

            /* Choice field specific enhancements */
            .ms-ChoiceGroup .tooltip-icon-inline,
            [data-control-name] .ms-ChoiceGroup .tooltip-icon-inline {
                margin: 0 0 0 8px !important;
                vertical-align: middle !important;
                align-self: flex-start !important;
                margin-top: 2px !important;
            }

            /* Responsive design */
            @media (max-width: 768px) {
                .standalone-tooltip-content {
                    max-width: calc(100vw - 32px) !important;
                    font-size: 13px !important;
                    padding: 12px 16px !important;
                }
                
                .tooltip-icon-base {
                    width: 14px !important;
                    height: 14px !important;
                    font-size: 9px !important;
                }
            }

            /* Dark theme support */
            @media (prefers-color-scheme: dark) {
                .tooltip-icon-base {
                    background: linear-gradient(135deg, #63b3ed 0%, #4299e1 100%) !important;
                    border-color: #3182ce !important;
                }
                
                .tooltip-icon-base:hover {
                    background: linear-gradient(135deg, #4299e1 0%, #3182ce 100%) !important;
                }
            }

            /* High contrast support */
            @media (prefers-contrast: high) {
                .tooltip-icon-base {
                    border: 2px solid ButtonText !important;
                    background: Highlight !important;
                    color: HighlightText !important;
                }
                
                .standalone-tooltip-content {
                    border: 2px solid WindowText !important;
                    background: Window !important;
                    color: WindowText !important;
                }
            }
        `;
        
        document.head.appendChild(style);
        console.log("Enhanced styles injected successfully");
    }

    private createTooltipElement(): void {
        this._tooltipElement = document.createElement('div');
        this._tooltipElement.className = 'standalone-tooltip-content';
        this._tooltipElement.id = `tooltip-content-${this._componentId}`;
        
        // Tooltip hover handlers
        this._tooltipElement.addEventListener('mouseenter', () => {
            this._isHovering = true;
            this.clearTimeouts();
        });
        
        this._tooltipElement.addEventListener('mouseleave', () => {
            this._isHovering = false;
            this.scheduleHide();
        });

        document.body.appendChild(this._tooltipElement);
        console.log("Tooltip element created and added to body");
    }

    private debugCurrentState(): void {
        console.log('=== TOOLTIP DEBUG STATE ===');
        console.log('Field Logical Name:', this.getStringParameter('fieldLogicalName'));
        console.log('Tooltip Content:', this.getTooltipContent());
        console.log('Is Hidden:', this._isHidden);
        console.log('Allow HTML:', this.getBooleanParameter('allowHtml', false));
        console.log('Tooltip Placement:', this.getStringParameter('tooltipPlacement', 'auto'));
    }

    private startPositioningStrategy(): void {
        console.log("Starting positioning strategy");
        
        // Attempt positioning with delays
        const delays = [100, 300, 800, 1500, 3000];
        delays.forEach((delay, index) => {
            setTimeout(() => {
                if (!this._positioned) {
                    console.log(`Positioning attempt ${index + 1} after ${delay}ms`);
                    this.attemptPositioning();
                }
            }, delay);
        });
        
        // DOM ready fallback
        if (document.readyState !== 'complete') {
            window.addEventListener('load', () => {
                setTimeout(() => {
                    if (!this._positioned) {
                        console.log("DOM loaded - attempting positioning");
                        this.attemptPositioning();
                    }
                }, 500);
            });
        }
        
        // Set up DOM observer
        this.setupDOMObserver();
    }

    private attemptPositioning(): void {
        if (this._positioned || this._positioningRetries >= this._maxRetries || this._isHidden) {
            return;
        }

        this._positioningRetries++;
        console.log(`Positioning attempt ${this._positioningRetries}`);

        // Strategy 1: Position by fieldLogicalName (BEST PRACTICE)
        const fieldLogicalName = this.getStringParameter('fieldLogicalName');
        if (fieldLogicalName && this.positionByFieldLogicalName(fieldLogicalName)) {
            this._positioned = true;
            return;
        }

        // Strategy 2: Enhanced Power Platform field search
        if (this.positionByPowerPlatformField()) {
            this._positioned = true;
            return;
        }

        // Strategy 3: Choice field specific positioning
        if (this.positionByChoiceField()) {
            this._positioned = true;
            return;
        }

        // Strategy 4: Find field by proximity to PCF container (NEW)
        if (this.positionByProximity()) {
            this._positioned = true;
            return;
        }

        console.log(`Positioning attempt ${this._positioningRetries} failed`);
    }

    private positionByProximity(): boolean {
        console.log("Attempting positioning by proximity to PCF container");
        
        try {
            // Get the container's parent to find nearby fields
            let searchRoot = this._container.parentElement;
            let depth = 0;
            
            // Walk up the DOM to find a good search root
            while (searchRoot && depth < 5) {
                if (searchRoot.querySelector('[data-control-name]') || 
                    searchRoot.classList.contains('ms-FieldGroup') ||
                    searchRoot.querySelector('.ms-TextField, .ms-Dropdown, .ms-ChoiceGroup')) {
                    break;
                }
                searchRoot = searchRoot.parentElement;
                depth++;
            }
            
            if (!searchRoot) return false;
            
            // Look for fields near the container
            const nearbyFields = searchRoot.querySelectorAll('[data-control-name], .ms-TextField-wrapper, .ms-Dropdown-container, .ms-ChoiceGroup');
            
            for (const field of Array.from(nearbyFields)) {
                const fieldElement = field as HTMLElement;
                if (this.isValidPositionTarget(fieldElement)) {
                    const controlName = fieldElement.getAttribute('data-control-name') || '';
                    console.log('Found nearby field by proximity:', fieldElement);
                    return this.attachToFieldElement(fieldElement, controlName);
                }
            }
            
        } catch (error) {
            console.warn('Error in proximity positioning:', error);
        }
        
        return false;
    }

    private positionByFieldLogicalName(logicalName: string): boolean {
        console.log(`Attempting positioning by logical name: ${logicalName}`);
        
        const selectors = [
            // Primary selectors for bound fields
            `[data-control-name="${logicalName}"]`,
            `div[data-lp-id="${logicalName}"]`,
            `input[data-control-name="${logicalName}"]`,
            
            // Choice field specific selectors
            `div[data-control-name="${logicalName}"] .ms-ChoiceGroup`,
            `div[data-control-name="${logicalName}"] .ms-ChoiceGroup-flexContainer`,
            
            // Text field selectors
            `div[data-control-name="${logicalName}"] .ms-TextField-wrapper`,
            `div[data-control-name="${logicalName}"] .ms-Dropdown-container`,
            
            // Fallback selectors
            `#${logicalName}`,
            `[name="${logicalName}"]`
        ];

        for (const selector of selectors) {
            try {
                const element = document.querySelector(selector) as HTMLElement;
                if (element && this.isValidPositionTarget(element)) {
                    console.log(`Found field by logical name with selector: ${selector}`, element);
                    return this.attachToFieldElement(element, logicalName);
                }
            } catch (error) {
                console.warn(`Error with selector ${selector}:`, error);
            }
        }

        return false;
    }

    private positionByPowerPlatformField(): boolean {
        console.log("Attempting Power Platform field positioning");
        
        const fieldSelectors = [
            // Choice field patterns
            '.ms-ChoiceGroup .ms-ChoiceGroup-flexContainer',
            '[data-control-name] .ms-ChoiceGroup',
            
            // Other field patterns
            '[data-control-name] .ms-TextField-wrapper',
            '[data-control-name] .ms-Dropdown-container',
            '[data-control-name] input[type="text"]',
        ];

        for (const selector of fieldSelectors) {
            try {
                const elements = document.querySelectorAll(selector);
                for (const element of Array.from(elements)) {
                    const fieldContainer = this.findFieldContainer(element as HTMLElement);
                    if (fieldContainer && this.isValidPositionTarget(fieldContainer)) {
                        console.log('Found Power Platform field:', fieldContainer);
                        const controlName = fieldContainer.getAttribute('data-control-name') || '';
                        return this.attachToFieldElement(fieldContainer, controlName);
                    }
                }
            } catch (error) {
                console.warn(`Error with selector ${selector}:`, error);
            }
        }

        return false;
    }

    private positionByChoiceField(): boolean {
        console.log("Attempting choice field specific positioning");
        
        const choiceSelectors = [
            '.ms-ChoiceGroup',
            '.ms-ChoiceGroup-flexContainer',
            '[role="radiogroup"]',
            'fieldset[data-control-name]'
        ];

        for (const selector of choiceSelectors) {
            try {
                const elements = document.querySelectorAll(selector);
                for (const element of Array.from(elements)) {
                    const fieldContainer = this.findFieldContainer(element as HTMLElement);
                    if (fieldContainer && this.isValidPositionTarget(fieldContainer)) {
                        console.log('Found choice field:', fieldContainer);
                        const controlName = fieldContainer.getAttribute('data-control-name') || '';
                        return this.attachToFieldElement(element as HTMLElement, controlName);
                    }
                }
            } catch (error) {
                console.warn(`Error with choice selector ${selector}:`, error);
            }
        }

        return false;
    }

    private findFieldContainer(element: HTMLElement): HTMLElement | null {
        let current = element;
        let depth = 0;
        const maxDepth = 8;

        while (current && depth < maxDepth) {
            if (current.hasAttribute('data-control-name') ||
                current.hasAttribute('data-lp-id')) {
                return current;
            }
            current = current.parentElement!;
            depth++;
        }

        return element;
    }

    private isValidPositionTarget(element: HTMLElement | null): boolean {
        if (!element || element === this._container || 
            element.contains(this._container) ||
            element.querySelector('.tooltip-icon-base')) {
            return false;
        }

        const rect = element.getBoundingClientRect();
        const isVisible = rect.width > 0 && rect.height > 0;
        
        return isVisible;
    }

    private attachToFieldElement(element: HTMLElement, controlName: string): boolean {
        try {
            // Create the icon
            const icon = this.createIcon();
            if (!icon) return false;

            // Find the label for this field to attach the icon there
            const fieldLabel = this.findFieldLabel(element, controlName);
            if (fieldLabel) {
                return this.attachToLabel(fieldLabel, icon, controlName);
            }

            // Fallback: Try to attach near the field without disrupting layout
            return this.attachNearField(element, icon, controlName);
            
        } catch (error) {
            console.warn('Failed to attach to field element:', error);
            return false;
        }
    }

    private findFieldLabel(element: HTMLElement, controlName: string): HTMLElement | null {
        // Strategy 1: Look for label with matching 'for' attribute
        const labelFor = document.querySelector(`label[for="${controlName}"]`) as HTMLElement;
        if (labelFor) return labelFor;

        // Strategy 2: Look for nearby label elements
        const fieldContainer = this.findFieldContainer(element);
        if (fieldContainer) {
            // Look for label within the same container
            const label = fieldContainer.querySelector('label') as HTMLElement;
            if (label) return label;

            // Look for div with role="group" or fieldset legend
            const groupLabel = fieldContainer.querySelector('legend, [role="group"] > div:first-child') as HTMLElement;
            if (groupLabel) return groupLabel;
        }

        // Strategy 3: Look for Power Platform specific label patterns
        const powerPlatformLabel = element.closest('[data-control-name]')?.querySelector('.ms-Label') as HTMLElement;
        if (powerPlatformLabel) return powerPlatformLabel;

        return null;
    }

    private attachToLabel(labelElement: HTMLElement, icon: HTMLElement, controlName: string): boolean {
        try {
            icon.className = 'tooltip-icon-base tooltip-icon-label';
            
            // Append icon directly to label without disrupting existing layout
            labelElement.appendChild(icon);
            
            this._iconElement = icon;
            this.setupIconEventListeners(icon);
            
            console.log(`Successfully attached tooltip to label for field: ${controlName}`, labelElement);
            return true;
            
        } catch (error) {
            console.warn('Failed to attach to label:', error);
            return false;
        }
    }

    private attachNearField(element: HTMLElement, icon: HTMLElement, controlName: string): boolean {
        try {
            // Determine positioning strategy based on field type
            const isChoiceGroup = element.classList.contains('ms-ChoiceGroup') ||
                                element.querySelector('.ms-ChoiceGroup') ||
                                element.getAttribute('role') === 'radiogroup';

            if (isChoiceGroup) {
                // For choice fields - position absolutely
                icon.className = 'tooltip-icon-base tooltip-icon-choice-absolute';
                
                // Make the choice group container relatively positioned
                const container = element.closest('[data-control-name]') as HTMLElement;
                if (container && container instanceof HTMLElement) {
                    container.style.position = 'relative';
                    container.appendChild(icon);
                }
                
            } else {
                // For text fields and dropdowns - position absolutely near the field
                icon.className = 'tooltip-icon-base tooltip-icon-field-absolute';
                
                const fieldWrapper = element.querySelector('.ms-TextField-wrapper, .ms-Dropdown-container');
                
                if (fieldWrapper && fieldWrapper instanceof HTMLElement) {
                    // Position relative to the field wrapper
                    fieldWrapper.style.position = 'relative';
                    
                    // Adjust for dropdowns which have their own arrow
                    if (element.querySelector('.ms-Dropdown')) {
                        icon.style.right = '32px !important';
                    }
                    
                    fieldWrapper.appendChild(icon);
                } else {
                    // Fallback: append to field container
                    const container = element.closest('[data-control-name]') as HTMLElement;
                    if (container) {
                        container.style.position = 'relative';
                        container.appendChild(icon);
                    }
                }
            }

            this._iconElement = icon;
            this.setupIconEventListeners(icon);
            
            console.log(`Successfully attached tooltip near field: ${controlName}`, element);
            return true;
            
        } catch (error) {
            console.warn('Failed to attach near field:', error);
            return false;
        }
    }

    private createIcon(): HTMLElement | null {
        const iconType = this.getStringParameter('iconType', 'info');
        const iconSize = this.getNumberParameter('iconSize', 10);

        const icon = document.createElement('span');
        icon.className = 'tooltip-icon-base';
        icon.textContent = this.getIconText(iconType);
        icon.setAttribute('tabindex', '0');
        icon.setAttribute('role', 'button');
        icon.setAttribute('aria-label', 'Show tooltip information');
        icon.style.fontSize = `${iconSize}px`;
        icon.title = '';  // Remove native tooltip to avoid conflicts

        return icon;
    }

    private getIconText(iconType: string): string {
        const icons: Record<string, string> = {
            'info': 'i',
            'question': '?',
            'warning': '⚠',
            'error': '!',
            'help': '?',
            'tip': '💡'
        };
        return icons[iconType] || 'i';
    }

    private setupDOMObserver(): void {
        const observer = new MutationObserver((mutations) => {
            if (!this._positioned && !this._isHidden) {
                if (this._observerTimeout !== undefined) {
                    clearTimeout(this._observerTimeout);
                }
                this._observerTimeout = window.setTimeout(() => {
                    this.attemptPositioning();
                }, 200);
            }
        });

        observer.observe(document.body, {
            childList: true,
            subtree: true,
            attributes: true,
            attributeFilter: ['class', 'data-control-name']
        });

        setTimeout(() => {
            observer.disconnect();
            if (this._observerTimeout !== undefined) {
                clearTimeout(this._observerTimeout);
            }
        }, 15000);
    }

    private setupIconEventListeners(iconElement: HTMLElement): void {
        const triggerType = this.getStringParameter('triggerType', 'hover');

        // Hover events
        if (triggerType === 'hover' || triggerType === 'both') {
            iconElement.addEventListener('mouseenter', () => {
                this._isHovering = true;
                this.scheduleShow();
            });
            
            iconElement.addEventListener('mouseleave', () => {
                this._isHovering = false;
                this.scheduleHide();
            });
        }

        // Click events
        if (triggerType === 'click' || triggerType === 'both') {
            iconElement.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.handleClick();
            });
        }

        // Keyboard events
        iconElement.addEventListener('keydown', (event) => {
            if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault();
                this.handleClick();
            }
            if (event.key === 'Escape' && this._isTooltipVisible) {
                this.hideTooltip();
            }
        });

        // Global events (set up once)
        if (!this._positioned) {
            document.addEventListener('click', this.handleDocumentClick);
            window.addEventListener('resize', this.handleResize);
            window.addEventListener('scroll', this.handleScroll, { passive: true });
        }
    }

    private scheduleShow(): void {
        this.clearTimeouts();
        const delay = this.getNumberParameter('showDelay', 300);
        this._showTimeout = window.setTimeout(() => {
            if (this._isHovering && !this._isTooltipVisible) {
                this.showTooltip();
            }
        }, delay);
    }

    private scheduleHide(): void {
        this.clearTimeouts();
        const delay = this.getNumberParameter('hideDelay', 100);
        this._hideTimeout = window.setTimeout(() => {
            if (!this._isHovering && this._isTooltipVisible) {
                this.hideTooltip();
            }
        }, delay);
    }

    private showTooltip(): void {
        if (!this._tooltipElement || !this._iconElement) return;

        console.log("Showing tooltip");
        this.updateTooltipContent();
        this.positionTooltip();
        
        requestAnimationFrame(() => {
            if (this._tooltipElement) {
                this._tooltipElement.classList.add('visible');
                this._isTooltipVisible = true;
            }
        });

        const autoHideDelay = this.getNumberParameter('autoHideDelay', 0);
        if (autoHideDelay > 0) {
            setTimeout(() => this.hideTooltip(), autoHideDelay);
        }
    }

    private hideTooltip(): void {
        if (!this._tooltipElement || !this._isTooltipVisible) return;
        
        console.log("Hiding tooltip");
        this._isTooltipVisible = false;
        this._tooltipElement.classList.remove('visible');
    }

    private updateTooltipContent(): void {
        if (!this._tooltipElement) return;

        const content = this.getTooltipContent() || 'No content provided';
        const allowHtml = this.getBooleanParameter('allowHtml', false);
        
        if (allowHtml) {
            // Clean and render HTML content properly
            this._tooltipElement.innerHTML = this.sanitizeHtml(content);
        } else {
            // Plain text content with line break support
            this._tooltipElement.innerHTML = this.escapeHtml(content).replace(/\n/g, '<br>');
        }
    }

    private positionTooltip(): void {
        if (!this._tooltipElement || !this._iconElement) return;

        // Get placement preference
        const placement = this.getStringParameter('tooltipPlacement', 'auto').toLowerCase();
        const iconRect = this._iconElement.getBoundingClientRect();
        const viewportWidth = window.innerWidth;
        const viewportHeight = window.innerHeight;
        const scrollX = window.scrollX;
        const scrollY = window.scrollY;
        
        // Temporarily show tooltip to get dimensions
        this._tooltipElement.style.visibility = 'hidden';
        this._tooltipElement.style.display = 'block';
        const tooltipRect = this._tooltipElement.getBoundingClientRect();
        this._tooltipElement.style.display = '';
        this._tooltipElement.style.visibility = '';
        
        let finalPlacement = placement;
        let top = 0;
        let left = 0;
        
        // Calculate positions for each placement option
        const positions = {
            bottom: {
                top: iconRect.bottom + scrollY + 12,
                left: iconRect.left + scrollX + (iconRect.width / 2) - (tooltipRect.width / 2),
                fits: iconRect.bottom + tooltipRect.height + 20 <= viewportHeight
            },
            top: {
                top: iconRect.top + scrollY - tooltipRect.height - 12,
                left: iconRect.left + scrollX + (iconRect.width / 2) - (tooltipRect.width / 2),
                fits: iconRect.top - tooltipRect.height - 20 >= 0
            },
            right: {
                top: iconRect.top + scrollY + (iconRect.height / 2) - (tooltipRect.height / 2),
                left: iconRect.right + scrollX + 12,
                fits: iconRect.right + tooltipRect.width + 20 <= viewportWidth
            },
            left: {
                top: iconRect.top + scrollY + (iconRect.height / 2) - (tooltipRect.height / 2),
                left: iconRect.left + scrollX - tooltipRect.width - 12,
                fits: iconRect.left - tooltipRect.width - 20 >= 0
            }
        };
        
        // Auto placement logic
        if (finalPlacement === 'auto') {
            const preferences = ['bottom', 'top', 'right', 'left'];
            finalPlacement = preferences.find(p => positions[p as keyof typeof positions].fits) || 'bottom';
        }
        
        // Use calculated position or fallback to bottom
        const position = positions[finalPlacement as keyof typeof positions] || positions.bottom;
        top = position.top;
        left = position.left;
        
        // Keep within viewport bounds
        const margin = 16;
        left = Math.max(margin, Math.min(left, viewportWidth - tooltipRect.width - margin));
        top = Math.max(margin, Math.min(top, viewportHeight - tooltipRect.height - margin));
        
        // Apply position
        this._tooltipElement.style.top = `${top}px`;
        this._tooltipElement.style.left = `${left}px`;
        this._tooltipElement.className = `standalone-tooltip-content ${finalPlacement}`;
        
        console.log(`Positioned tooltip: ${finalPlacement} at (${left}, ${top})`);
    }

    private handleClick(): void {
        console.log("Icon clicked");
        
        const redirectUrl = this.getStringParameter('redirectUrl');
        
        if (redirectUrl?.trim()) {
            const sanitizedUrl = this.sanitizeUrl(redirectUrl.trim());
            if (sanitizedUrl) {
                const openInNewTab = this.getBooleanParameter('openInNewTab', true);
                if (openInNewTab) {
                    window.open(sanitizedUrl, '_blank', 'noopener,noreferrer');
                } else {
                    window.location.href = sanitizedUrl;
                }
                return;
            }
        }
        
        // Toggle tooltip visibility
        if (this._isTooltipVisible) {
            this.hideTooltip();
        } else {
            this.showTooltip();
        }
    }

    private handleDocumentClick = (event: MouseEvent): void => {
        if (!this._isTooltipVisible || !event.target) return;
        
        const target = event.target as HTMLElement;
        if (this._tooltipElement && this._iconElement &&
            !this._tooltipElement.contains(target) && 
            !this._iconElement.contains(target)) {
            this.hideTooltip();
        }
    }

    private handleResize = (): void => {
        if (this._isTooltipVisible) {
            this.positionTooltip();
        }
    }

    private handleScroll = (): void => {
        if (this._isTooltipVisible) {
            this.positionTooltip();
        }
    }

    private clearTimeouts(): void {
        if (this._hideTimeout !== undefined) {
            clearTimeout(this._hideTimeout);
            this._hideTimeout = undefined;
        }
        if (this._showTimeout !== undefined) {
            clearTimeout(this._showTimeout);
            this._showTimeout = undefined;
        }
    }

    private createFallbackIcon(): void {
        // Only create fallback if not hidden
        if (this._isHidden) return;
        
        console.log("Creating fallback icon - positioning failed");
        
        // Create a temporary container that's visible but minimal
        const fallbackContainer = document.createElement('div');
        fallbackContainer.style.cssText = `
            position: fixed !important;
            top: 10px !important;
            right: 10px !important;
            z-index: 999999 !important;
            background: rgba(255, 0, 0, 0.1) !important;
            border: 2px dashed #ff0000 !important;
            padding: 4px !important;
            border-radius: 4px !important;
            font-size: 11px !important;
            color: #ff0000 !important;
            font-family: monospace !important;
        `;
        
        const fallbackIcon = document.createElement('span');
        fallbackIcon.className = 'tooltip-icon-base';
        fallbackIcon.style.cssText = `
            display: inline-flex !important;
            background: linear-gradient(135deg, #e53e3e 0%, #c53030 100%) !important;
            color: white !important;
            font-weight: bold !important;
            border-color: #c53030 !important;
        `;
        fallbackIcon.textContent = '!';
        fallbackIcon.title = 'Tooltip positioning failed - check fieldLogicalName parameter';
        fallbackIcon.setAttribute('tabindex', '0');
        fallbackIcon.setAttribute('role', 'button');
        fallbackIcon.setAttribute('aria-label', 'Show tooltip information (positioning failed)');
        
        fallbackContainer.appendChild(fallbackIcon);
        fallbackContainer.appendChild(document.createTextNode(' Fallback'));
        document.body.appendChild(fallbackContainer);
        
        this._iconElement = fallbackIcon;
        this.setupIconEventListeners(fallbackIcon);
        this._positioned = true;
        
        // Remove fallback after 10 seconds
        setTimeout(() => {
            if (fallbackContainer.parentNode) {
                fallbackContainer.parentNode.removeChild(fallbackContainer);
            }
        }, 10000);
        
        console.log("Fallback icon created - will be removed after 10 seconds");
    }

    // Parameter helper methods
    private getTooltipContent(): string {
        return this._context?.parameters?.tooltipContent?.formatted || 
               this._context?.parameters?.tooltipContent?.raw || 
               'Default tooltip content';
    }

    private getStringParameter(name: keyof IInputs, defaultValue = ""): string {
        const value = this._context?.parameters?.[name]?.raw;
        return value !== undefined ? String(value) : defaultValue;
    }

    private getNumberParameter(name: keyof IInputs, defaultValue = 0): number {
        const value = this._context?.parameters?.[name]?.raw;
        const num = Number(value);
        return !isNaN(num) ? num : defaultValue;
    }

    private getBooleanParameter(name: keyof IInputs, defaultValue = false): boolean {
        const value = this._context?.parameters?.[name]?.raw;
        if (typeof value === 'boolean') return value;
        if (typeof value === 'string') return value.toLowerCase() === 'true';
        return defaultValue;
    }

    // Security and content helpers
    private escapeHtml(text: string): string {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    private sanitizeHtml(html: string): string {
        // Create a clean container
        const temp = document.createElement('div');
        temp.innerHTML = html;
        
        // Remove only the most dangerous elements
        const dangerousElements = temp.querySelectorAll('script, style, link, iframe, object, embed, form, input, button');
        dangerousElements.forEach(el => el.remove());
        
        // Allow safe HTML tags including img for icons
        const allowedTags = [
            'b', 'strong', 'i', 'em', 'u', 'br', 'p', 'span', 'div', 
            'ul', 'ol', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
            'img', 'a', 'small', 'code', 'pre', 'blockquote'
        ];
        
        const elements = temp.querySelectorAll('*');
        
        for (let i = 0; i < elements.length; i++) {
            const element = elements[i] as HTMLElement;
            if (!allowedTags.includes(element.tagName.toLowerCase())) {
                // Replace with span but keep content
                const span = document.createElement('span');
                span.innerHTML = element.innerHTML;
                element.parentNode?.replaceChild(span, element);
            } else {
                // Clean attributes for allowed elements
                this.cleanElementAttributes(element);
            }
        }
        
        return temp.innerHTML;
    }

    private cleanElementAttributes(element: HTMLElement): void {
        const tagName = element.tagName.toLowerCase();
        const attrs = element.attributes;
        
        // Define allowed attributes per tag
        const allowedAttributes: Record<string, string[]> = {
            'img': ['src', 'alt', 'title', 'width', 'height', 'style'],
            'a': ['href', 'title', 'target'],
            'span': ['class', 'style'],
            'div': ['class', 'style'],
            'p': ['class', 'style'],
            'strong': ['class', 'style'],
            'b': ['class', 'style'],
            'em': ['class', 'style'],
            'i': ['class', 'style'],
            'u': ['class', 'style'],
            'small': ['class', 'style'],
            'code': ['class', 'style'],
            'pre': ['class', 'style'],
            'blockquote': ['class', 'style'],
            'ul': ['class', 'style'],
            'ol': ['class', 'style'],
            'li': ['class', 'style'],
            'h1': ['class', 'style'],
            'h2': ['class', 'style'],
            'h3': ['class', 'style'],
            'h4': ['class', 'style'],
            'h5': ['class', 'style'],
            'h6': ['class', 'style']
        };
        
        const allowed = allowedAttributes[tagName] || [];
        
        // Remove disallowed attributes
        for (let j = attrs.length - 1; j >= 0; j--) {
            const attr = attrs[j];
            const attrName = attr.name.toLowerCase();
            
            // Always remove event handlers and dangerous attributes
            if (attrName.startsWith('on') || 
                attrName === 'action' ||
                attrName === 'formaction' ||
                attrName === 'javascript:') {
                element.removeAttribute(attr.name);
            }
            // Remove attributes not in allowed list
            else if (!allowed.includes(attrName)) {
                element.removeAttribute(attr.name);
            }
            // Validate specific attributes
            else if (attrName === 'src' && tagName === 'img') {
                // Allow data URLs for inline images and common image URLs
                const src = attr.value.toLowerCase();
                if (!src.startsWith('data:image/') && 
                    !src.startsWith('http://') && 
                    !src.startsWith('https://') &&
                    !src.startsWith('/')) {
                    element.removeAttribute(attr.name);
                }
            }
            else if (attrName === 'href' && tagName === 'a') {
                // Sanitize href values
                const href = attr.value.toLowerCase();
                if (href.startsWith('javascript:') || 
                    href.startsWith('vbscript:') ||
                    href.startsWith('data:')) {
                    element.removeAttribute(attr.name);
                }
            }
        }
        
        // Add safe styling for images to prevent layout issues
        if (tagName === 'img') {
            const currentStyle = element.getAttribute('style') || '';
            if (!currentStyle.includes('max-width')) {
                element.setAttribute('style', `${currentStyle} max-width: 100%; height: auto;`.trim());
            }
        }
    }

    private sanitizeUrl(url: string): string | null {
        try {
            const urlObj = new URL(url);
            const allowedProtocols = ['http:', 'https:', 'mailto:', 'tel:'];
            return allowedProtocols.includes(urlObj.protocol) ? urlObj.toString() : null;
        } catch {
            return null;
        }
    }

    // PCF Lifecycle Methods
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        
        // Check if visibility state changed
        const newHiddenState = this.getBooleanParameter('isHidden', false);
        if (newHiddenState !== this._isHidden) {
            this._isHidden = newHiddenState;
            
            if (this._isHidden) {
                // Hide everything
                this.hideContainerCompletely();
                if (this._iconElement && this._iconElement.parentElement) {
                    this._iconElement.remove();
                    this._iconElement = null;
                }
                if (this._tooltipElement) {
                    this.hideTooltip();
                }
                this._positioned = false;
            } else {
                // Show and try to reposition
                this.hideContainerCompletely(); // Keep container hidden
                this._positioned = false;
                this._positioningRetries = 0;
                this.startPositioningStrategy();
            }
        }
        
        // Update icon appearance if visible
        if (this._iconElement && !this._isHidden) {
            const iconType = this.getStringParameter('iconType', 'info');
            const iconSize = this.getNumberParameter('iconSize', 10);
            this._iconElement.textContent = this.getIconText(iconType);
            this._iconElement.style.fontSize = `${iconSize}px`;
        }
        
        // Update tooltip content if visible
        if (this._isTooltipVisible) {
            this.updateTooltipContent();
            this.positionTooltip();
        }
        
        // Continue positioning attempts if not yet positioned and not hidden
        if (!this._positioned && !this._isHidden && this._positioningRetries < this._maxRetries) {
            setTimeout(() => this.attemptPositioning(), 100);
        }
        
        console.log("UpdateView completed - positioned:", this._positioned, "hidden:", this._isHidden);
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
    console.log("Destroying tooltip component with aggressive cleanup:", this._componentId);
    
    // Your existing cleanup code...
    this.clearTimeouts();
    
    if (this._observerTimeout !== undefined) {
        clearTimeout(this._observerTimeout);
    }
    
    document.removeEventListener('click', this.handleDocumentClick);
    window.removeEventListener('resize', this.handleResize);
    window.removeEventListener('scroll', this.handleScroll);
    
    if (this._tooltipElement && this._tooltipElement.parentNode) {
        this._tooltipElement.parentNode.removeChild(this._tooltipElement);
    }
    
    if (this._iconElement && this._iconElement.parentElement !== this._container) {
        this._iconElement.remove();
    }
    
    // Clean up aggressively hidden elements
    const forcedHidden = document.querySelectorAll('.pcf-tooltip-force-hidden, [data-pcf-hidden="true"]');
    forcedHidden.forEach(element => {
        const el = element as HTMLElement;
        el.classList.remove('pcf-tooltip-force-hidden');
        el.removeAttribute('data-pcf-hidden');
        // Optionally restore display (be careful not to break other components)
        // el.style.display = '';
    });
    
    // Clean up graveyard and styles
    const otherTooltips = document.querySelectorAll('[id*="tooltip-content-"]');
    if (otherTooltips.length <= 1) {
        const graveyard = document.getElementById('pcf-hidden-graveyard');
        if (graveyard) {
            graveyard.remove();
        }
        
        const aggressiveStyles = document.getElementById('pcf-tooltip-aggressive-hiding');
        if (aggressiveStyles) {
            aggressiveStyles.remove();
        }
    }
    
    // Reset state
    this._positioned = false;
    this._isTooltipVisible = false;
    this._iconElement = null;
    this._tooltipElement = null;
    
    console.log("Aggressive cleanup completed");
}
}