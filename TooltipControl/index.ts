import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class StandaloneTooltipComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _tooltipElement: HTMLDivElement | null = null;
    private _iconElement: HTMLElement | null = null;
    
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
            
            this._isHidden = this.getBooleanParameter('isHidden', false);
            
            this.eliminateSpaceCompletely();
            
            if (this._isHidden) {
                console.log("Component is set to hidden");
                return;
            }
            
            this.injectStyles();
            this.createTooltipElement();
            this.startPositioningStrategy();
            
            console.log("Tooltip component initialized");
        } catch (error) {
            console.error("Error during initialization:", error);
            this.eliminateSpaceCompletely();
            
            if (!this._isHidden) {
                this.createFallbackIcon();
            }
        }
    }

    private eliminateSpaceCompletely(): void {
        console.log("Eliminating PCF container space");
        
        if (this._container) {
            this._container.style.cssText = `
                display: none !important;
                position: absolute !important;
                left: -99999px !important;
                top: -99999px !important;
                width: 0 !important;
                height: 0 !important;
                margin: 0 !important;
                padding: 0 !important;
                border: none !important;
                visibility: hidden !important;
                opacity: 0 !important;
                z-index: -999999 !important;
                pointer-events: none !important;
            `;
            this._container.innerHTML = '';
        }
        
        this.eliminateParentSpacing();
        this.injectNuclearCSS();
    }

    private eliminateParentSpacing(): void {
        if (!this._container) return;
        
        let current: HTMLElement = this._container;
        let depth = 0;
        const maxDepth = 5;
        
        while (current && depth < maxDepth) {
            const parent = current.parentElement;
            if (!parent) break;
            
            const seemsPCFRelated = 
                parent.querySelector('div') === this._container ||
                parent.children.length === 1 ||
                parent.classList.contains('ms-FormField') ||
                parent.hasAttribute('data-control-name');
                
            if (seemsPCFRelated) {
                parent.style.cssText = `
                    display: none !important;
                    height: 0 !important;
                    margin: 0 !important;
                    padding: 0 !important;
                    border: none !important;
                    visibility: hidden !important;
                `;
                parent.classList.add('pcf-tooltip-eliminated');
            }
            
            current = parent;
            depth++;
        }
    }

    private injectNuclearCSS(): void {
        const styleId = 'pcf-tooltip-nuclear-elimination';
        if (document.getElementById(styleId)) return;

        const style = document.createElement('style');
        style.id = styleId;
        style.textContent = `
            .pcf-tooltip-eliminated {
                display: none !important;
                height: 0 !important;
                margin: 0 !important;
                padding: 0 !important;
                border: none !important;
                visibility: hidden !important;
            }
        `;
        document.head.appendChild(style);
    }

    private injectStyles(): void {
        const styleId = 'standalone-tooltip-styles';
        if (document.getElementById(styleId)) return;

        const iconColor = this.getStringParameter('iconColor', '#ffffff');
        const iconBackgroundColor = this.getStringParameter('iconBackgroundColor', '#4299e1');
        const iconBorderColor = this.getStringParameter('iconBorderColor', '#3182ce');
        const tooltipBgColor = this.getStringParameter('tooltipBackgroundColor', '#323130');
        const tooltipTextColor = this.getStringParameter('tooltipTextColor', '#ffffff');
        const maxWidth = this.getNumberParameter('maxWidth', 350);
        const minWidth = this.getNumberParameter('minWidth', 200);
        const borderRadius = this.getNumberParameter('tooltipBorderRadius', 8);
        const fontSize = this.getNumberParameter('tooltipFontSize', 13);
        const iconSize = this.getNumberParameter('iconSize', 16);

        const style = document.createElement('style');
        style.id = styleId;
        style.textContent = `
            .tooltip-icon-base {
                display: inline-flex !important;
                align-items: center !important;
                justify-content: center !important;
                width: ${iconSize}px !important;
                height: ${iconSize}px !important;
                cursor: pointer !important;
                border-radius: 50% !important;
                background: ${iconBackgroundColor} !important;
                border: 1px solid ${iconBorderColor} !important;
                font-size: ${Math.max(8, iconSize - 6)}px !important;
                font-weight: 700 !important;
                color: ${iconColor} !important;
                vertical-align: middle !important;
                flex-shrink: 0 !important;
                transition: all 0.2s ease !important;
                z-index: 10000 !important;
                position: relative !important;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1) !important;
                font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, sans-serif !important;
                user-select: none !important;
                -webkit-user-select: none !important;
                -moz-user-select: none !important;
                -ms-user-select: none !important;
                pointer-events: auto !important;
                touch-action: manipulation !important;
                -webkit-tap-highlight-color: transparent !important;
                text-align: center !important;
                line-height: 1 !important;
                white-space: nowrap !important;
                overflow: hidden !important;
            }
            
            .tooltip-icon-base:hover {
                transform: scale(1.1) !important;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15) !important;
            }
            
            .tooltip-icon-base:focus {
                outline: 2px solid ${iconBorderColor} !important;
                outline-offset: 2px !important;
            }

            .tooltip-icon-inline {
                margin: 0 0 0 8px !important;
            }

            .tooltip-icon-absolute {
                position: absolute !important;
                right: 8px !important;
                top: 50% !important;
                transform: translateY(-50%) !important;
                margin: 0 !important;
                z-index: 10000 !important;
            }

            .tooltip-icon-label {
                margin: 0 0 0 6px !important;
                vertical-align: baseline !important;
            }

            .standalone-tooltip-content {
                position: fixed !important;
                background: ${tooltipBgColor} !important;
                background-color: ${tooltipBgColor} !important;
                color: ${tooltipTextColor} !important;
                padding: 12px 16px !important;
                border-radius: ${borderRadius}px !important;
                font-size: ${fontSize}px !important;
                line-height: 1.4 !important;
                max-width: ${maxWidth}px !important;
                min-width: ${minWidth}px !important;
                width: auto !important;
                height: auto !important;
                z-index: ${this.getNumberParameter('zIndex', 999999)} !important;
                box-shadow: 0 8px 16px rgba(0, 0, 0, 0.15) !important;
                border: 1px solid rgba(255, 255, 255, 0.1) !important;
                visibility: hidden !important;
                opacity: 0 !important;
                transform: translateY(-8px) scale(0.95) !important;
                transition: all 0.2s ease !important;
                pointer-events: none !important;
                word-wrap: break-word !important;
                font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, sans-serif !important;
                overflow: hidden !important;
                ${this.getBooleanParameter('enableBackdropFilter', false) ? 'backdrop-filter: blur(8px) !important;' : ''}
            }
            
            .standalone-tooltip-content,
            .standalone-tooltip-content::before,
            .standalone-tooltip-content::after {
                background: ${tooltipBgColor} !important;
                background-color: ${tooltipBgColor} !important;
            }
            
            .standalone-tooltip-content.visible {
                visibility: visible !important;
                opacity: 1 !important;
                transform: translateY(0) scale(1) !important;
                pointer-events: auto !important;
            }
            
            .standalone-tooltip-content.visible.click-mode {
                pointer-events: none !important;
            }

            ${this.getBooleanParameter('enableArrow', true) ? this.getArrowStyles(tooltipBgColor) : ''}

            .standalone-tooltip-content, 
            .standalone-tooltip-content *:not(img) {
                color: ${tooltipTextColor} !important;
                background: transparent !important;
                background-color: transparent !important;
            }

            .standalone-tooltip-content {
                background: ${tooltipBgColor} !important;
                background-color: ${tooltipBgColor} !important;
            }

            .standalone-tooltip-content p {
                margin: 0 0 8px 0 !important;
                color: ${tooltipTextColor} !important;
                background: transparent !important;
            }

            .standalone-tooltip-content p:last-child {
                margin-bottom: 0 !important;
            }

            .standalone-tooltip-content h1,
            .standalone-tooltip-content h2,
            .standalone-tooltip-content h3,
            .standalone-tooltip-content h4,
            .standalone-tooltip-content h5,
            .standalone-tooltip-content h6 {
                color: ${tooltipTextColor} !important;
                margin: 0 0 8px 0 !important;
                font-weight: 600 !important;
                background: transparent !important;
                background-color: transparent !important;
            }

            .standalone-tooltip-content strong,
            .standalone-tooltip-content b {
                font-weight: 600 !important;
                color: ${tooltipTextColor} !important;
                background: transparent !important;
            }

            .standalone-tooltip-content em,
            .standalone-tooltip-content i {
                font-style: italic !important;
                color: ${tooltipTextColor} !important;
                background: transparent !important;
            }

            .standalone-tooltip-content ul,
            .standalone-tooltip-content ol {
                margin: 8px 0 !important;
                padding-left: 20px !important;
                color: ${tooltipTextColor} !important;
                background: transparent !important;
            }

            .standalone-tooltip-content li {
                margin: 4px 0 !important;
                color: ${tooltipTextColor} !important;
                background: transparent !important;
            }

            .standalone-tooltip-content a {
                color: #60a5fa !important;
                text-decoration: underline !important;
                background: transparent !important;
            }

            .standalone-tooltip-content a:hover {
                color: #93c5fd !important;
            }

            .standalone-tooltip-content div {
                color: ${tooltipTextColor} !important;
                background: transparent !important;
                background-color: transparent !important;
            }

            .standalone-tooltip-content img {
                max-width: 50px !important;
                max-height: 50px !important;
                width: auto !important;
                height: auto !important;
                border-radius: 4px !important;
                margin: 4px 0 !important;
                display: inline-block !important;
                vertical-align: middle !important;
                object-fit: contain !important;
            }

            @media (max-width: 768px) {
                .standalone-tooltip-content {
                    max-width: calc(100vw - 32px) !important;
                    font-size: ${Math.max(12, fontSize - 1)}px !important;
                    padding: 10px 14px !important;
                }
                
                .tooltip-icon-base {
                    width: ${Math.max(14, iconSize - 2)}px !important;
                    height: ${Math.max(14, iconSize - 2)}px !important;
                    font-size: ${Math.max(7, iconSize - 7)}px !important;
                }
            }
        `;
        
        document.head.appendChild(style);
    }

    private getArrowStyles(bgColor: string): string {
        return `
            .standalone-tooltip-content::before {
                content: '' !important;
                position: absolute !important;
                border: 8px solid transparent !important;
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
        `;
    }

    private createTooltipElement(): void {
        this._tooltipElement = document.createElement('div');
        this._tooltipElement.className = 'standalone-tooltip-content';
        this._tooltipElement.id = `tooltip-content-${this._componentId}`;
        
        this._tooltipElement.addEventListener('mouseenter', () => {
            this._isHovering = true;
            this.clearTimeouts();
        });
        
        this._tooltipElement.addEventListener('mouseleave', () => {
            this._isHovering = false;
            this.scheduleHide();
        });

        document.body.appendChild(this._tooltipElement);
    }

    private startPositioningStrategy(): void {
        const delays = [100, 300, 800, 1500];
        delays.forEach((delay, index) => {
            setTimeout(() => {
                if (!this._positioned) {
                    this.attemptPositioning();
                }
            }, delay);
        });
        
        this.setupDOMObserver();
    }

    private attemptPositioning(): void {
        if (this._positioned || this._positioningRetries >= this._maxRetries || this._isHidden) {
            return;
        }

        this._positioningRetries++;
        console.log(`Positioning attempt ${this._positioningRetries}`);

        const fieldLogicalName = this.getStringParameter('fieldLogicalName');
        if (fieldLogicalName && this.positionByFieldLogicalName(fieldLogicalName)) {
            this._positioned = true;
            return;
        }

        if (this.positionByPowerPlatformField()) {
            this._positioned = true;
            return;
        }

        if (this._positioningRetries >= 3) {
            console.log(`Creating fallback after ${this._positioningRetries} attempts`);
            this.createFallbackIcon();
        }

        console.log(`Positioning attempt ${this._positioningRetries} failed`);
    }

    private positionByFieldLogicalName(logicalName: string): boolean {
        const selectors = [
            `[data-control-name="${logicalName}"]`,
            `div[data-lp-id="${logicalName}"]`,
            `input[data-control-name="${logicalName}"]`,
            `div[data-control-name="${logicalName}"] .ms-ChoiceGroup`,
            `div[data-control-name="${logicalName}"] .ms-TextField-wrapper`,
            `div[data-control-name="${logicalName}"] .ms-Dropdown-container`,
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
        const fieldSelectors = [
            '.ms-ChoiceGroup .ms-ChoiceGroup-flexContainer',
            '[data-control-name] .ms-ChoiceGroup',
            '[data-control-name] .ms-TextField-wrapper',
            '[data-control-name] .ms-Dropdown-container',
        ];

        for (const selector of fieldSelectors) {
            try {
                const elements = document.querySelectorAll(selector);
                for (const element of Array.from(elements)) {
                    const fieldContainer = this.findFieldContainer(element as HTMLElement);
                    if (fieldContainer && this.isValidPositionTarget(fieldContainer)) {
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

    private findFieldContainer(element: HTMLElement): HTMLElement | null {
        let current = element;
        let depth = 0;
        const maxDepth = 6;

        while (current && depth < maxDepth) {
            if (current.hasAttribute('data-control-name') || current.hasAttribute('data-lp-id')) {
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
        return rect.width > 0 && rect.height > 0;
    }

    private attachToFieldElement(element: HTMLElement, controlName: string): boolean {
        try {
            const icon = this.createIcon();
            if (!icon) return false;

            const fieldLabel = this.findFieldLabel(element, controlName);
            if (fieldLabel) {
                return this.attachToLabel(fieldLabel, icon, controlName);
            }

            return this.attachNearField(element, icon, controlName);
            
        } catch (error) {
            console.warn('Failed to attach to field element:', error);
            return false;
        }
    }

    private findFieldLabel(element: HTMLElement, controlName: string): HTMLElement | null {
        const labelFor = document.querySelector(`label[for="${controlName}"]`) as HTMLElement;
        if (labelFor) return labelFor;

        const fieldContainer = this.findFieldContainer(element);
        if (fieldContainer) {
            const label = fieldContainer.querySelector('label') as HTMLElement;
            if (label) return label;

            const groupLabel = fieldContainer.querySelector('legend, [role="group"] > div:first-child') as HTMLElement;
            if (groupLabel) return groupLabel;
        }

        return null;
    }

    private attachToLabel(labelElement: HTMLElement, icon: HTMLElement, controlName: string): boolean {
        try {
            icon.className = 'tooltip-icon-base tooltip-icon-label';
            labelElement.appendChild(icon);
            
            this._iconElement = icon;
            this.setupIconEventListeners(icon);
            
            console.log(`Successfully attached tooltip to label for field: ${controlName}`);
            return true;
            
        } catch (error) {
            console.warn('Failed to attach to label:', error);
            return false;
        }
    }

    private attachNearField(element: HTMLElement, icon: HTMLElement, controlName: string): boolean {
        try {
            const isChoiceGroup = element.classList.contains('ms-ChoiceGroup') ||
                                element.querySelector('.ms-ChoiceGroup') ||
                                element.getAttribute('role') === 'radiogroup';

            if (isChoiceGroup) {
                icon.className = 'tooltip-icon-base tooltip-icon-absolute';
                const container = element.closest('[data-control-name]') as HTMLElement;
                if (container) {
                    container.style.position = 'relative';
                    container.appendChild(icon);
                }
            } else {
                icon.className = 'tooltip-icon-base tooltip-icon-absolute';
                const fieldWrapper = element.querySelector('.ms-TextField-wrapper, .ms-Dropdown-container');
                
                if (fieldWrapper) {
                    (fieldWrapper as HTMLElement).style.position = 'relative';
                    fieldWrapper.appendChild(icon);
                } else {
                    const container = element.closest('[data-control-name]') as HTMLElement;
                    if (container) {
                        container.style.position = 'relative';
                        container.appendChild(icon);
                    }
                }
            }

            this._iconElement = icon;
            this.setupIconEventListeners(icon);
            
            console.log(`Successfully attached tooltip near field: ${controlName}`);
            return true;
            
        } catch (error) {
            console.warn('Failed to attach near field:', error);
            return false;
        }
    }

    private createIcon(): HTMLElement | null {
        const iconType = this.getStringParameter('iconType', 'info');
        const customIcon = this.getStringParameter('customIcon', '');
        const iconSize = this.getNumberParameter('iconSize', 16);
        const ariaLabel = this.getStringParameter('ariaLabel', 'Show tooltip information');

        const icon = document.createElement('span');
        icon.className = 'tooltip-icon-base';
        
        let iconContent = customIcon;
        if (!iconContent || iconContent.trim() === '' || iconContent === 'null' || iconContent === 'undefined') {
            iconContent = this.getIconText(iconType);
        }
        
        icon.innerHTML = '';
        icon.textContent = iconContent;
        
        icon.setAttribute('tabindex', '0');
        icon.setAttribute('role', 'button');
        icon.setAttribute('aria-label', ariaLabel);
        icon.style.fontSize = `${Math.max(8, iconSize - 6)}px`;
        icon.title = '';
        
        icon.style.userSelect = 'none';
        icon.style.pointerEvents = 'auto';
        icon.style.cursor = 'pointer';
        icon.style.touchAction = 'manipulation';
        icon.style.setProperty('-webkit-tap-highlight-color', 'transparent');

        icon.style.position = 'relative';
        icon.style.zIndex = '10001';

        return icon;
    }

    private getIconText(iconType: string): string {
        const icons: Record<string, string> = {
            'info': 'i',
            'question': '?',
            'warning': '⚠',
            'error': '!',
            'help': '?'
        };
        return icons[iconType] || 'i';
    }

    private setupDOMObserver(): void {
        const observer = new MutationObserver(() => {
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
        }, 10000);
    }

    private setupIconEventListeners(iconElement: HTMLElement): void {
        const triggerType = this.getStringParameter('triggerType', 'hover');
        const enableKeyboard = this.getBooleanParameter('enableKeyboardNavigation', true);

        iconElement.style.zIndex = '10001';
        iconElement.style.position = 'relative';
        iconElement.style.pointerEvents = 'auto';
        iconElement.style.userSelect = 'none';
        iconElement.style.webkitUserSelect = 'none';
        iconElement.style.cursor = 'pointer';
        iconElement.style.touchAction = 'manipulation';

        iconElement.addEventListener('selectstart', (e) => {
            e.preventDefault();
            return false;
        });

        iconElement.addEventListener('contextmenu', (e) => {
            e.preventDefault();
            return false;
        });

        const handleClickAction = (e: Event) => {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            
            if (e.target === iconElement) {
                this.handleClick();
            }
        };

        iconElement.addEventListener('click', handleClickAction, { capture: true, passive: false });
        iconElement.addEventListener('mouseup', handleClickAction, { capture: true, passive: false });
        iconElement.addEventListener('touchend', handleClickAction, { capture: true, passive: false });

        iconElement.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
        }, { capture: true, passive: false });

        iconElement.addEventListener('touchstart', (e) => {
            e.preventDefault();
            e.stopPropagation();
        }, { capture: true, passive: false });

        if (triggerType === 'hover' || triggerType === 'both') {
            iconElement.addEventListener('mouseenter', (e) => {
                e.stopPropagation();
                this._isHovering = true;
                this.scheduleShow();
            });
            
            iconElement.addEventListener('mouseleave', (e) => {
                e.stopPropagation();
                this._isHovering = false;
                this.scheduleHide();
            });
        }

        if (enableKeyboard) {
            iconElement.addEventListener('keydown', (event) => {
                if (event.key === 'Enter' || event.key === ' ') {
                    event.preventDefault();
                    event.stopPropagation();
                    this.handleClick();
                }
                if (event.key === 'Escape' && this._isTooltipVisible) {
                    this.hideTooltip();
                }
            });
        }

        document.addEventListener('click', this.handleDocumentClick);
        window.addEventListener('resize', this.handleResize);
        window.addEventListener('scroll', this.handleScroll, { passive: true });
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

        this.updateTooltipContent();
        this.positionTooltip();
        
        const triggerType = this.getStringParameter('triggerType', 'hover');
        if (triggerType === 'click' || triggerType === 'both') {
            this._tooltipElement.classList.add('click-mode');
        }
        
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
        
        this._isTooltipVisible = false;
        this._tooltipElement.classList.remove('visible', 'click-mode');
    }

    private updateTooltipContent(): void {
        if (!this._tooltipElement) return;

        let content = this.getTooltipContent() || 'No content provided';
        const allowHtml = this.getBooleanParameter('allowHtml', false);
        const preserveLineBreaks = this.getBooleanParameter('preserveLineBreaks', true);
        
        if (allowHtml) {
            content = this.cleanHtmlContent(content);
            this._tooltipElement.innerHTML = this.sanitizeHtml(content);
        } else {
            const processedContent = preserveLineBreaks ? 
                this.escapeHtml(content).replace(/\n/g, '<br>') : 
                this.escapeHtml(content);
            this._tooltipElement.innerHTML = processedContent;
        }
    }

    private cleanHtmlContent(html: string): string {
        let cleaned = html.replace(/style\s*=\s*["'][^"']*["']/gi, '');
        
        cleaned = cleaned.replace(/font-family\s*:\s*[^;]+;?/gi, '');
        
        cleaned = cleaned.replace(/color\s*:\s*[^;]+;?/gi, '');
        
        cleaned = cleaned.replace(/background[^;]*:\s*[^;]+;?/gi, '');
        
        cleaned = cleaned.replace(/margin\s*:\s*[^;]+;?/gi, '');
        cleaned = cleaned.replace(/padding\s*:\s*[^;]+;?/gi, '');
        
        cleaned = cleaned.replace(/<img([^>]*)\s+width\s*=\s*["'][^"']*["']([^>]*)>/gi, '<img$1$2>');
        cleaned = cleaned.replace(/<img([^>]*)\s+height\s*=\s*["'][^"']*["']([^>]*)>/gi, '<img$1$2>');
        
        cleaned = cleaned.replace(/style\s*=\s*["']\s*["']/gi, '');
        
        return cleaned;
    }

    private positionTooltip(): void {
        if (!this._tooltipElement || !this._iconElement) return;

        const placement = this.getStringParameter('tooltipPlacement', 'auto').toLowerCase();
        const iconRect = this._iconElement.getBoundingClientRect();
        const viewportWidth = window.innerWidth;
        const viewportHeight = window.innerHeight;
        const scrollX = window.scrollX;
        const scrollY = window.scrollY;
        
        this._tooltipElement.style.visibility = 'hidden';
        this._tooltipElement.style.display = 'block';
        const tooltipRect = this._tooltipElement.getBoundingClientRect();
        this._tooltipElement.style.display = '';
        this._tooltipElement.style.visibility = '';
        
        let finalPlacement = placement;
        let top = 0;
        let left = 0;
        
        const spacing = 16;
        
        const positions = {
            bottom: {
                top: iconRect.bottom + scrollY + spacing,
                left: iconRect.left + scrollX + (iconRect.width / 2) - (tooltipRect.width / 2),
                fits: iconRect.bottom + tooltipRect.height + spacing + 20 <= viewportHeight
            },
            top: {
                top: iconRect.top + scrollY - tooltipRect.height - spacing,
                left: iconRect.left + scrollX + (iconRect.width / 2) - (tooltipRect.width / 2),
                fits: iconRect.top - tooltipRect.height - spacing - 20 >= 0
            },
            right: {
                top: iconRect.top + scrollY + (iconRect.height / 2) - (tooltipRect.height / 2),
                left: iconRect.right + scrollX + spacing,
                fits: iconRect.right + tooltipRect.width + spacing + 20 <= viewportWidth
            },
            left: {
                top: iconRect.top + scrollY + (iconRect.height / 2) - (tooltipRect.height / 2),
                left: iconRect.left + scrollX - tooltipRect.width - spacing,
                fits: iconRect.left - tooltipRect.width - spacing - 20 >= 0
            }
        };
        
        if (finalPlacement === 'auto') {
            const preferences = ['bottom', 'top', 'right', 'left'];
            finalPlacement = preferences.find(p => positions[p as keyof typeof positions].fits) || 'bottom';
        }
        
        const position = positions[finalPlacement as keyof typeof positions] || positions.bottom;
        top = position.top;
        left = position.left;
        
        const margin = 20;
        left = Math.max(margin, Math.min(left, viewportWidth - tooltipRect.width - margin));
        top = Math.max(margin, Math.min(top, viewportHeight - tooltipRect.height - margin));
        
        const finalTooltipRect = {
            left: left,
            right: left + tooltipRect.width,
            top: top,
            bottom: top + tooltipRect.height
        };
        
        const iconBounds = {
            left: iconRect.left + scrollX,
            right: iconRect.right + scrollX,
            top: iconRect.top + scrollY,
            bottom: iconRect.bottom + scrollY
        };
        
        if (this.rectsOverlap(finalTooltipRect, iconBounds)) {
            if (finalPlacement === 'bottom' || finalPlacement === 'top') {
                if (iconRect.right + tooltipRect.width + spacing + 20 <= viewportWidth) {
                    left = iconRect.right + scrollX + spacing;
                    finalPlacement = 'right';
                } else if (iconRect.left - tooltipRect.width - spacing - 20 >= 0) {
                    left = iconRect.left + scrollX - tooltipRect.width - spacing;
                    finalPlacement = 'left';
                }
            } else {
                if (iconRect.bottom + tooltipRect.height + spacing + 20 <= viewportHeight) {
                    top = iconRect.bottom + scrollY + spacing;
                    finalPlacement = 'bottom';
                } else if (iconRect.top - tooltipRect.height - spacing - 20 >= 0) {
                    top = iconRect.top + scrollY - tooltipRect.height - spacing;
                    finalPlacement = 'top';
                }
            }
        }
        
        this._tooltipElement.style.top = `${top}px`;
        this._tooltipElement.style.left = `${left}px`;
        this._tooltipElement.className = `standalone-tooltip-content ${finalPlacement}`;
    }

    private rectsOverlap(rect1: any, rect2: any): boolean {
        return !(rect1.right < rect2.left || 
                rect2.right < rect1.left || 
                rect1.bottom < rect2.top || 
                rect2.bottom < rect1.top);
    }

    private handleClick(): void {
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
        if (this._isHidden) return;
        
        console.log("Creating fallback icon - positioning failed");
        
        const fallbackContainer = document.createElement('div');
        fallbackContainer.style.cssText = `
            position: fixed !important;
            top: 20px !important;
            right: 20px !important;
            z-index: 10002 !important;
            background: rgba(255, 255, 255, 0.95) !important;
            border: 2px solid #4299e1 !important;
            padding: 8px 12px !important;
            border-radius: 8px !important;
            font-size: 12px !important;
            color: #1a365d !important;
            font-family: "Segoe UI", sans-serif !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15) !important;
            max-width: 250px !important;
        `;
        
        const fallbackIcon = this.createIcon();
        if (fallbackIcon) {
            fallbackIcon.style.cssText += `
                display: inline-flex !important;
                margin-right: 8px !important;
                background: #4299e1 !important;
                color: white !important;
                font-weight: bold !important;
                border-color: #3182ce !important;
                z-index: 10003 !important;
                position: relative !important;
            `;
            
            fallbackContainer.appendChild(fallbackIcon);
            fallbackContainer.appendChild(document.createTextNode('Tooltip (Fallback Mode)'));
            document.body.appendChild(fallbackContainer);
            
            this._iconElement = fallbackIcon;
            this.setupIconEventListeners(fallbackIcon);
            this._positioned = true;
            
            setTimeout(() => {
                if (fallbackContainer.parentNode) {
                    fallbackContainer.parentNode.removeChild(fallbackContainer);
                }
            }, 15000);
        }
    }

    private getTooltipContent(): string {
        const content = this._context?.parameters?.tooltipContent?.raw;
        return content !== null && content !== undefined ? String(content) : 'Default tooltip content';
    }

    private getStringParameter(name: keyof IInputs, defaultValue = ""): string {
        const value = this._context?.parameters?.[name]?.raw;
        
        const stringValue = value !== null && value !== undefined ? String(value) : '';
        
        if (stringValue === 'null' || stringValue === 'undefined' || stringValue === '') {
            return defaultValue;
        }
        
        return stringValue;
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

    private escapeHtml(text: string): string {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    private sanitizeHtml(html: string): string {
        const temp = document.createElement('div');
        temp.innerHTML = html;
        
        const dangerousElements = temp.querySelectorAll('script, style, link, iframe, object, embed, form, input, button');
        dangerousElements.forEach(el => el.remove());
        
        const allowedTags = ['b', 'strong', 'i', 'em', 'u', 'br', 'p', 'span', 'div', 'ul', 'ol', 'li', 'a', 'small', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'img'];
        const elements = temp.querySelectorAll('*');
        
        for (let i = 0; i < elements.length; i++) {
            const element = elements[i] as HTMLElement;
            if (!allowedTags.includes(element.tagName.toLowerCase())) {
                const span = document.createElement('span');
                span.innerHTML = element.innerHTML;
                element.parentNode?.replaceChild(span, element);
            } else {
                this.cleanElementAttributes(element);
            }
        }
        
        return temp.innerHTML;
    }

    private cleanElementAttributes(element: HTMLElement): void {
        const attrs = element.attributes;
        for (let j = attrs.length - 1; j >= 0; j--) {
            const attr = attrs[j];
            const attrName = attr.name.toLowerCase();
            
            if (attrName.startsWith('on') || 
                attrName === 'action' ||
                attrName === 'formaction' ||
                attrName === 'style') {
                element.removeAttribute(attr.name);
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

    private debugParameters(): void {
        console.log('=== Tooltip Parameter Debug ===');
        
        const iconTypeParam = this._context?.parameters?.iconType;
        console.log('iconType:', {
            raw: iconTypeParam?.raw,
            final: this.getStringParameter('iconType', 'info')
        });
        
        const customIconParam = this._context?.parameters?.customIcon;
        console.log('customIcon:', {
            raw: customIconParam?.raw,
            final: this.getStringParameter('customIcon', '')
        });
        
        console.log('================================');
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        
        const newHiddenState = this.getBooleanParameter('isHidden', false);
        if (newHiddenState !== this._isHidden) {
            this._isHidden = newHiddenState;
            
            if (this._isHidden) {
                this.eliminateSpaceCompletely();
                if (this._iconElement && this._iconElement.parentElement) {
                    this._iconElement.remove();
                    this._iconElement = null;
                }
                if (this._tooltipElement) {
                    this.hideTooltip();
                }
                this._positioned = false;
            } else {
                this.eliminateSpaceCompletely();
                this._positioned = false;
                this._positioningRetries = 0;
                this.startPositioningStrategy();
            }
        }
        
        if (this._iconElement && !this._isHidden) {
            const iconType = this.getStringParameter('iconType', 'info');
            const customIcon = this.getStringParameter('customIcon', '');
            const iconSize = this.getNumberParameter('iconSize', 16);
            
            let iconContent = customIcon;
            if (!iconContent || iconContent.trim() === '' || iconContent === 'null' || iconContent === 'undefined') {
                iconContent = this.getIconText(iconType);
            }
            
            this._iconElement.textContent = iconContent;
            this._iconElement.style.fontSize = `${Math.max(8, iconSize - 6)}px`;
        }
        
        if (this._isTooltipVisible) {
            this.updateTooltipContent();
            this.positionTooltip();
        }
        
        if (!this._positioned && !this._isHidden && this._positioningRetries < this._maxRetries) {
            setTimeout(() => this.attemptPositioning(), 100);
        }
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        console.log("Destroying tooltip component:", this._componentId);
        
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
        
        const eliminatedElements = document.querySelectorAll('.pcf-tooltip-eliminated');
        eliminatedElements.forEach(element => {
            const el = element as HTMLElement;
            el.classList.remove('pcf-tooltip-eliminated');
            el.style.cssText = '';
        });
        
        this._positioned = false;
        this._isTooltipVisible = false;
        this._iconElement = null;
        this._tooltipElement = null;
    }
}