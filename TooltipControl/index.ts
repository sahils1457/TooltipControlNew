import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class TooltipControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _tooltipElement: HTMLDivElement;
    private _iconElement: HTMLElement;
    private _inputElement: HTMLInputElement | HTMLTextAreaElement;
    private _inputContainer: HTMLDivElement;
    private _styleElement: HTMLStyleElement;
    private _showTooltipBound: () => void;
    private _hideTooltipBound: () => void;
    private _clickHandlerBound: (event: MouseEvent) => void;
    private _resizeHandlerBound: () => void;
    private _inputChangeBound: (event: Event) => void;
    private _documentClickBound: (event: MouseEvent) => void;
    private _hideTimeout: number | null = null;
    private _showTimeout: number | null = null;
    private _currentValue = '';
    private _isTooltipVisible = false;
    private _observer: ResizeObserver | null = null;
    private _isHovering = false; // Track hover state

    constructor() {
        this._showTooltipBound = this.showTooltip.bind(this);
        this._hideTooltipBound = this.hideTooltip.bind(this);
        this._clickHandlerBound = this.handleIconClick.bind(this);
        this._resizeHandlerBound = this.handleResize.bind(this);
        this._inputChangeBound = this.handleInputChange.bind(this);
        this._documentClickBound = this.handleDocumentClick.bind(this);
    }

    public init(
        context: ComponentFramework.Context<IInputs>, 
        notifyOutputChanged: () => void, 
        state: ComponentFramework.Dictionary, 
        container: HTMLDivElement
    ): void {
        if (!container) {
            console.error("Container is undefined or null");
            return;
        }

        this._context = context;
        this._container = container;
        this._notifyOutputChanged = notifyOutputChanged;

        this.injectGlobalStyles();
        this.createInputWithTooltip();
        this.setupEventListeners();
        this.setupResizeObserver();
    }

    private createInputWithTooltip(): void {
        if (!this._container) {
            console.error("Container is not available");
            return;
        }

        this._container.innerHTML = '';
        this._container.className = 'tooltip-control-wrapper';

        // Create the main input container with native Power Platform styling
        this._inputContainer = document.createElement('div');
        this._inputContainer.className = 'input-with-tooltip-container';

        // Create the input field
        this.createInputField();

        // Create the tooltip icon
        this.createTooltipIcon();

        // Create the tooltip element
        this.createTooltipElement();

        // Assemble the components
        this._inputContainer.appendChild(this._inputElement);
        this._inputContainer.appendChild(this._iconElement);
        this._container.appendChild(this._inputContainer);
        
        // Append tooltip to document body for proper z-index
        document.body.appendChild(this._tooltipElement);

        // Apply styling
        this.applyNativeFieldStyles();
    }

    private createInputField(): void {
        const isMultiline = this.getBooleanParameter('isMultiline', false);
        const maxLength = this.getNumberParameter('maxLength', 0);
        
        if (isMultiline) {
            this._inputElement = document.createElement('textarea') as HTMLTextAreaElement;
            const rows = this.getNumberParameter('rows', 3);
            (this._inputElement as HTMLTextAreaElement).rows = rows;
        } else {
            this._inputElement = document.createElement('input') as HTMLInputElement;
            (this._inputElement as HTMLInputElement).type = 'text';
        }

        // Set common properties
        this._inputElement.className = 'tooltip-native-input';
        
        if (maxLength > 0) {
            this._inputElement.maxLength = maxLength;
        }

        // Set placeholder
        const placeholder = this.getStringParameter('placeholder', '');
        if (placeholder) {
            this._inputElement.placeholder = placeholder;
        }

        // Set initial value
        this._currentValue = this.getStringParameter('dummyField', '');
        this._inputElement.value = this._currentValue;

        // Set readonly if specified
        const isReadonly = this.getBooleanParameter('isReadonly', false);
        if (isReadonly) {
            this._inputElement.readOnly = true;
        }

        // Add event listeners
        this._inputElement.addEventListener('input', this._inputChangeBound);
        this._inputElement.addEventListener('change', this._inputChangeBound);
        this._inputElement.addEventListener('blur', this._inputChangeBound);
    }

    private createTooltipIcon(): void {
        this._iconElement = document.createElement('span');
        this._iconElement.className = 'tooltip-native-icon';
        
        const iconColor = this.getStringParameter('iconColor', '#605e5c');
        const iconSize = this.getNumberParameter('iconSize', 16);
        
        this._iconElement.style.cssText = `
            color: ${iconColor};
            font-size: ${iconSize}px;
        `;
        
        this.updateIconContent();

        // FIXED: Improved hover handling
        this._iconElement.addEventListener('mouseenter', this.handleIconMouseEnter.bind(this));
        this._iconElement.addEventListener('mouseleave', this.handleIconMouseLeave.bind(this));
        this._iconElement.addEventListener('click', this._clickHandlerBound);
        this._iconElement.addEventListener('focus', this._showTooltipBound);
        this._iconElement.addEventListener('blur', this._hideTooltipBound);

        // Accessibility attributes
        this._iconElement.setAttribute('tabindex', '0');
        this._iconElement.setAttribute('role', 'button');
        this._iconElement.setAttribute('aria-label', 'Show help information');
        this._iconElement.setAttribute('aria-describedby', 'tooltip-content');
    }

    // FIXED: Simplified and more reliable hover handling
    private handleIconMouseEnter(): void {
        console.log("Icon mouse enter");
        this._isHovering = true;
        
        // Clear any existing timeouts
        this.clearAllTimeouts();
        
        // Show tooltip immediately if not visible, or with slight delay
        const delay = this._isTooltipVisible ? 0 : 150;
        
        this._showTimeout = window.setTimeout(() => {
            if (this._isHovering) { // Double-check we're still hovering
                this.showTooltip();
            }
            this._showTimeout = null;
        }, delay);
    }

    private handleIconMouseLeave(): void {
        console.log("Icon mouse leave");
        this._isHovering = false;
        
        // Clear show timeout
        if (this._showTimeout) {
            clearTimeout(this._showTimeout);
            this._showTimeout = null;
        }
        
        // Set hide timeout
        this._hideTimeout = window.setTimeout(() => {
            if (!this._isHovering) { // Only hide if we're truly not hovering
                this.hideTooltip();
            }
            this._hideTimeout = null;
        }, 300);
    }

    private createTooltipElement(): void {
        this._tooltipElement = document.createElement('div');
        this._tooltipElement.className = 'tooltip-native-content';
        this._tooltipElement.id = 'tooltip-content';
        
        // FIXED: Better tooltip hover handling
        this._tooltipElement.addEventListener('mouseenter', () => {
            console.log("Tooltip mouse enter");
            this._isHovering = true;
            this.clearAllTimeouts();
        });
        
        this._tooltipElement.addEventListener('mouseleave', () => {
            console.log("Tooltip mouse leave");
            this._isHovering = false;
            
            this._hideTimeout = window.setTimeout(() => {
                if (!this._isHovering) {
                    this.hideTooltip();
                }
                this._hideTimeout = null;
            }, 200);
        });

        this.updateTooltipContent();
    }

    // FIXED: Helper method to clear all timeouts
    private clearAllTimeouts(): void {
        if (this._hideTimeout) {
            clearTimeout(this._hideTimeout);
            this._hideTimeout = null;
        }
        
        if (this._showTimeout) {
            clearTimeout(this._showTimeout);
            this._showTimeout = null;
        }
    }

    private applyNativeFieldStyles(): void {
        if (!this._inputContainer || !this._inputElement || !this._iconElement) return;

        const inputWidth = this.getStringParameter('inputWidth', '100%');
        const inputHeight = this.getStringParameter('inputHeight', 'auto');
        
        // Main wrapper - match Power Platform container behavior
        this._container.style.cssText = `
            position: relative;
            display: block;
            width: 100%;
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        `;
        
        // Input container - relative positioning for icon placement
        this._inputContainer.style.cssText = `
            position: relative;
            display: block;
            width: ${inputWidth};
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        `;

        // Native Power Platform input styling
        const isMultiline = this._inputElement.tagName.toLowerCase() === 'textarea';
        const baseHeight = isMultiline ? '80px' : '32px';
        const finalHeight = inputHeight === 'auto' ? baseHeight : inputHeight;
        
        this._inputElement.style.cssText = `
            width: 100% !important;
            height: ${finalHeight} !important;
            min-height: ${finalHeight} !important;
            padding: ${isMultiline ? '8px 36px 8px 12px' : '6px 36px 6px 12px'} !important;
            border: 1px solid #8a8886 !important;
            border-radius: 2px !important;
            font-size: 14px !important;
            font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, "Roboto", "Helvetica Neue", sans-serif !important;
            line-height: 20px !important;
            background-color: #ffffff !important;
            color: #323130 !important;
            outline: none !important;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.23, 1) !important;
            box-sizing: border-box !important;
            resize: ${isMultiline ? 'vertical' : 'none'} !important;
            overflow: ${isMultiline ? 'auto' : 'hidden'} !important;
            word-wrap: break-word !important;
            position: relative !important;
            z-index: 1 !important;
        `;

        // Icon positioning and styling
        this._iconElement.style.cssText = `
            position: absolute !important;
            right: 8px !important;
            top: ${isMultiline ? '12px' : '50%'} !important;
            transform: ${isMultiline ? 'none' : 'translateY(-50%)'} !important;
            z-index: 10 !important;
            width: 20px !important;
            height: 20px !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            cursor: help !important;
            border-radius: 3px !important;
            transition: all 0.2s ease !important;
            background: rgba(255, 255, 255, 0.9) !important;
            border: 1px solid transparent !important;
            user-select: none !important;
            pointer-events: auto !important;
            font-weight: 600 !important;
            text-align: center !important;
            line-height: 1 !important;
            box-sizing: border-box !important;
        `;

        this.setupInputFocusHandlers();
    }

    private setupInputFocusHandlers(): void {
        if (!this._inputElement) return;

        const focusHandler = () => {
            this._inputElement.style.setProperty('border-color', '#0078d4', 'important');
            this._inputElement.style.setProperty('box-shadow', '0 0 0 1px #0078d4', 'important');
            
            // Highlight the icon when input is focused
            if (this._iconElement) {
                this._iconElement.style.setProperty('background', '#f3f2f1', 'important');
                this._iconElement.style.setProperty('border-color', '#0078d4', 'important');
            }
        };

        const blurHandler = () => {
            this._inputElement.style.setProperty('border-color', '#8a8886', 'important');
            this._inputElement.style.setProperty('box-shadow', 'none', 'important');
            
            // Reset icon highlight when input loses focus
            if (this._iconElement && !this._isTooltipVisible) {
                this._iconElement.style.setProperty('background', 'rgba(255, 255, 255, 0.9)', 'important');
                this._iconElement.style.setProperty('border-color', 'transparent', 'important');
            }
        };

        this._inputElement.addEventListener('focus', focusHandler);
        this._inputElement.addEventListener('blur', blurHandler);
    }

    private handleInputChange(event: Event): void {
        if (!this._inputElement) return;

        const newValue = this._inputElement.value;
        if (newValue !== this._currentValue) {
            this._currentValue = newValue;
            this._notifyOutputChanged();
        }
    }

    private injectGlobalStyles(): void {
        if (this._styleElement && this._styleElement.parentNode) {
            this._styleElement.parentNode.removeChild(this._styleElement);
        }

        this._styleElement = document.createElement('style');
        this._styleElement.id = `tooltip-native-styles-${Date.now()}`;
        
        const bgColor = this.getStringParameter('tooltipBackgroundColor', '#323130');
        const textColor = this.getStringParameter('tooltipTextColor', '#ffffff');
        const linkColor = this.getLinkColor(textColor);
        
        this._styleElement.textContent = this.generateNativeStyleContent(bgColor, textColor, linkColor);
        document.head.appendChild(this._styleElement);
    }

    private generateNativeStyleContent(bgColor: string, textColor: string, linkColor: string): string {
        return `
            /* Main wrapper */
            .tooltip-control-wrapper {
                position: relative !important;
                display: block !important;
                width: 100% !important;
                margin: 0 !important;
                padding: 0 !important;
                background: transparent !important;
                box-sizing: border-box !important;
            }
            
            /* Input container */
            .input-with-tooltip-container {
                position: relative !important;
                display: block !important;
                width: 100% !important;
                margin: 0 !important;
                padding: 0 !important;
                background: transparent !important;
                box-sizing: border-box !important;
            }
            
            /* Native Power Platform input styling */
            .tooltip-native-input {
                width: 100% !important;
                padding: 6px 36px 6px 12px !important;
                border: 1px solid #8a8886 !important;
                border-radius: 2px !important;
                font-size: 14px !important;
                font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, "Roboto", "Helvetica Neue", sans-serif !important;
                line-height: 20px !important;
                background-color: #ffffff !important;
                color: #323130 !important;
                outline: none !important;
                transition: all 0.2s cubic-bezier(0.4, 0, 0.23, 1) !important;
                box-sizing: border-box !important;
                position: relative !important;
                z-index: 1 !important;
                height: 32px !important;
                min-height: 32px !important;
            }
            
            /* Textarea specific adjustments */
            textarea.tooltip-native-input {
                height: 80px !important;
                min-height: 80px !important;
                padding: 8px 36px 8px 12px !important;
                resize: vertical !important;
                overflow-y: auto !important;
                line-height: 1.5 !important;
            }
            
            /* Focus state - matches native Power Platform */
            .tooltip-native-input:focus {
                border-color: #0078d4 !important;
                box-shadow: 0 0 0 1px #0078d4 !important;
                background-color: #ffffff !important;
            }
            
            /* Hover state */
            .tooltip-native-input:hover:not(:focus) {
                border-color: #605e5c !important;
            }
            
            /* Error state */
            .tooltip-native-input.error {
                border-color: #d13438 !important;
                box-shadow: 0 0 0 1px #d13438 !important;
            }
            
            /* Disabled/readonly state */
            .tooltip-native-input:disabled,
            .tooltip-native-input:read-only {
                background-color: #f3f2f1 !important;
                color: #a19f9d !important;
                border-color: #edebe9 !important;
                cursor: not-allowed !important;
            }
            
            /* Placeholder styling */
            .tooltip-native-input::placeholder {
                color: #a19f9d !important;
                font-style: normal !important;
                opacity: 1 !important;
            }
            
            /* Native tooltip icon */
            .tooltip-native-icon {
                position: absolute !important;
                right: 8px !important;
                top: 50% !important;
                transform: translateY(-50%) !important;
                width: 20px !important;
                height: 20px !important;
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
                cursor: help !important;
                border-radius: 3px !important;
                background: rgba(255, 255, 255, 0.9) !important;
                border: 1px solid transparent !important;
                transition: all 0.2s ease !important;
                z-index: 10 !important;
                user-select: none !important;
                pointer-events: auto !important;
                font-weight: 600 !important;
                text-shadow: 0 1px 2px rgba(0, 0, 0, 0.1) !important;
                color: #605e5c !important;
                box-sizing: border-box !important;
            }
            
            /* Textarea icon positioning */
            textarea.tooltip-native-input + .tooltip-native-icon {
                top: 12px !important;
                transform: none !important;
            }
            
            /* Icon states */
            .tooltip-native-icon:hover {
                background: #f3f2f1 !important;
                border-color: #d2d0ce !important;
                color: #0078d4 !important;
                transform: translateY(-50%) scale(1.05) !important;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1) !important;
            }
            
            textarea.tooltip-native-input + .tooltip-native-icon:hover {
                transform: scale(1.05) !important;
            }
            
            .tooltip-native-icon:active {
                transform: translateY(-50%) scale(0.95) !important;
                background: #edebe9 !important;
            }
            
            textarea.tooltip-native-input + .tooltip-native-icon:active {
                transform: scale(0.95) !important;
            }
            
            .tooltip-native-icon:focus {
                outline: 2px solid #0078d4 !important;
                outline-offset: 1px !important;
                background: #f3f2f1 !important;
                border-color: #0078d4 !important;
            }
            
            /* Highlighted/active icon state */
            .tooltip-native-icon.highlighted {
                background: #deecf9 !important;
                border-color: #0078d4 !important;
                color: #0078d4 !important;
                box-shadow: 0 2px 8px rgba(0, 120, 212, 0.2) !important;
            }
            
            /* Clickable icon styles */
            .tooltip-native-icon.clickable {
                cursor: pointer !important;
            }
            
            .tooltip-native-icon.clickable:hover {
                background: #e1dfdd !important;
                color: #005a9e !important;
                border-color: #0078d4 !important;
            }
            
            /* FIXED: Improved tooltip positioning and overflow handling */
            .tooltip-native-content {
                position: fixed !important;
                background: ${bgColor} !important;
                color: ${textColor} !important;
                padding: 12px 16px !important;
                border-radius: 4px !important;
                font-size: 14px !important;
                line-height: 1.5 !important;
                max-width: min(320px, calc(100vw - 40px)) !important;
                min-width: min(200px, calc(100vw - 40px)) !important;
                z-index: 2147483647 !important;
                box-shadow: 0 8px 32px rgba(0, 0, 0, 0.14), 0 4px 8px rgba(0, 0, 0, 0.12) !important;
                border: 1px solid rgba(255, 255, 255, 0.1) !important;
                visibility: hidden !important;
                opacity: 0 !important;
                transform: translateY(-4px) !important;
                transition: all 0.2s cubic-bezier(0.4, 0, 0.23, 1) !important;
                pointer-events: none !important;
                word-wrap: break-word !important;
                overflow-wrap: break-word !important;
                font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, "Roboto", "Helvetica Neue", sans-serif !important;
                box-sizing: border-box !important;
                white-space: normal !important;
                text-align: left !important;
                filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.1)) !important;
                
                /* Ensure tooltip stays within viewport */
                max-height: calc(100vh - 40px) !important;
                overflow-y: auto !important;
                
                /* Better positioning for responsive layouts */
                left: 0 !important;
                top: 0 !important;
            }
            
            /* Visible tooltip with smooth animation */
            .tooltip-native-content.visible {
                visibility: visible !important;
                opacity: 1 !important;
                transform: translateY(0) !important;
                pointer-events: auto !important;
            }
            
            /* Tooltip arrow */
            .tooltip-native-content::before {
                content: '' !important;
                position: absolute !important;
                width: 0 !important;
                height: 0 !important;
                border: 6px solid transparent !important;
                border-bottom-color: ${bgColor} !important;
                top: -12px !important;
                left: 50% !important;
                transform: translateX(-50%) !important;
                filter: drop-shadow(0 -2px 2px rgba(0, 0, 0, 0.1)) !important;
            }
            
            /* HTML content styling within tooltip */
            .tooltip-native-content h1, .tooltip-native-content h2, .tooltip-native-content h3, 
            .tooltip-native-content h4, .tooltip-native-content h5, .tooltip-native-content h6 {
                margin: 8px 0 4px 0 !important;
                padding: 0 !important;
                font-weight: 600 !important;
                color: ${textColor} !important;
                line-height: 1.3 !important;
            }
            
            .tooltip-native-content h1 { font-size: 18px !important; }
            .tooltip-native-content h2 { font-size: 16px !important; }
            .tooltip-native-content h3 { font-size: 15px !important; }
            .tooltip-native-content h4, .tooltip-native-content h5, .tooltip-native-content h6 { 
                font-size: 14px !important; 
            }
            
            .tooltip-native-content p {
                margin: 6px 0 !important;
                padding: 0 !important;
                color: ${textColor} !important;
                line-height: 1.5 !important;
            }
            
            .tooltip-native-content ul, .tooltip-native-content ol {
                margin: 8px 0 !important;
                padding-left: 20px !important;
                color: ${textColor} !important;
            }
            
            .tooltip-native-content li {
                margin: 2px 0 !important;
                color: ${textColor} !important;
            }
            
            .tooltip-native-content a {
                color: ${linkColor} !important;
                text-decoration: underline !important;
                transition: opacity 0.2s ease !important;
            }
            
            .tooltip-native-content a:hover {
                opacity: 0.8 !important;
            }
            
            .tooltip-native-content strong, .tooltip-native-content b {
                font-weight: 600 !important;
                color: ${textColor} !important;
            }
            
            .tooltip-native-content em, .tooltip-native-content i {
                font-style: italic !important;
                color: ${textColor} !important;
            }
            
            .tooltip-native-content > *:first-child { margin-top: 0 !important; }
            .tooltip-native-content > *:last-child { margin-bottom: 0 !important; }

            /* FIXED: Enhanced responsive design */
            @media screen and (max-width: 768px) {
                .tooltip-native-content {
                    max-width: calc(100vw - 20px) !important;
                    min-width: calc(100vw - 40px) !important;
                    font-size: 13px !important;
                    padding: 10px 14px !important;
                    left: 10px !important;
                    right: 10px !important;
                    width: auto !important;
                }
                
                .tooltip-native-input {
                    font-size: 16px !important; /* Prevent zoom on iOS */
                }
            }
            
            /* Power Platform integration fixes */
            [data-control-name] .tooltip-control-wrapper,
            .customcontrol-container .tooltip-control-wrapper,
            .form-component .tooltip-control-wrapper {
                position: relative !important;
                width: 100% !important;
                height: auto !important;
            }
            
            /* Ensure proper stacking in Power Platform */
            .ms-Fabric .tooltip-native-content {
                z-index: 2147483647 !important;
            }
            
            /* Override any conflicting Power Platform styles */
            .tooltip-native-input {
                min-width: unset !important;
                max-width: unset !important;
            }

            /* FIXED: Better handling of multi-column layouts */
            @media screen and (min-width: 769px) {
                .tooltip-native-content {
                    max-width: min(400px, 45vw) !important;
                }
            }
            
            /* For very narrow containers (like sidebars) */
            @container (max-width: 300px) {
                .tooltip-native-content {
                    max-width: calc(100vw - 20px) !important;
                    min-width: 200px !important;
                }
            }
        `;
    }

    private setupEventListeners(): void {
        // Document click handler to hide tooltip when clicking outside
        document.addEventListener('click', this._documentClickBound);
        
        // Window events
        window.addEventListener('resize', this._resizeHandlerBound);
        window.addEventListener('scroll', this._resizeHandlerBound, true);
    }

    private handleDocumentClick(event: MouseEvent): void {
        if (!this._isTooltipVisible) return;
        
        const target = event.target as Element;
        if (target && 
            !this._tooltipElement.contains(target) && 
            !this._iconElement.contains(target)) {
            this._isHovering = false;
            this.hideTooltip();
        }
    }

    private setupResizeObserver(): void {
        if ('ResizeObserver' in window) {
            this._observer = new ResizeObserver(() => {
                if (this._isTooltipVisible) {
                    this.positionTooltip();
                }
            });
            
            if (this._container) {
                this._observer.observe(this._container);
            }
        }
    }

    private getLinkColor(textColor: string): string {
        const isDark = this.isColorDark(textColor);
        return isDark ? '#4fc3f7' : '#0078d4';
    }

    private isColorDark(color: string): boolean {
        if (color.startsWith('#')) {
            const hex = color.slice(1);
            const r = parseInt(hex.substr(0, 2), 16);
            const g = parseInt(hex.substr(2, 2), 16);
            const b = parseInt(hex.substr(4, 2), 16);
            const brightness = (r * 299 + g * 587 + b * 114) / 1000;
            return brightness < 128;
        }
        return color.toLowerCase().includes('dark') || color.toLowerCase() === '#000000';
    }

    private getParameterValue(parameterName: keyof IInputs): string | number | boolean | null {
        try {
            if (this._context?.parameters?.[parameterName]) {
                const param = this._context.parameters[parameterName];
                return param.raw !== undefined ? param.raw : null;
            }
        } catch (error) {
            console.warn(`Error getting parameter ${String(parameterName)}:`, error);
        }
        return null;
    }

    private getStringParameter(parameterName: keyof IInputs, defaultValue = ""): string {
        const value = this.getParameterValue(parameterName);
        if (value === null || value === undefined) return defaultValue;
        return String(value);
    }

    private getNumberParameter(parameterName: keyof IInputs, defaultValue = 0): number {
        const value = this.getParameterValue(parameterName);
        if (value === null || value === undefined) return defaultValue;
        const numValue = Number(value);
        return isNaN(numValue) ? defaultValue : numValue;
    }

    private getBooleanParameter(parameterName: keyof IInputs, defaultValue = false): boolean {
        const value = this.getParameterValue(parameterName);
        if (value === null || value === undefined) return defaultValue;
        if (typeof value === 'boolean') return value;
        if (typeof value === 'string') {
            const lowerValue = value.toLowerCase();
            return lowerValue === 'true' || lowerValue === '1' || lowerValue === 'yes';
        }
        return Boolean(value);
    }

    private updateIconContent(): void {
        if (!this._iconElement) return;

        const iconType = this.getStringParameter('iconType', 'info');
        this._iconElement.textContent = this.getDefaultIcon(iconType);
        
        // Update clickability
        const redirectUrl = this.getStringParameter('redirectUrl');
        const hasRedirect = Boolean(redirectUrl && redirectUrl.trim());
        
        if (hasRedirect) {
            this._iconElement.classList.add('clickable');
            this._iconElement.style.cursor = 'pointer';
            this._iconElement.setAttribute('aria-label', 'Show help and click to navigate');
        } else {
            this._iconElement.classList.remove('clickable');
            this._iconElement.style.cursor = 'help';
            this._iconElement.setAttribute('aria-label', 'Show help information');
        }
    }

    private updateTooltipContent(): void {
        if (!this._tooltipElement) return;

        let tooltipContent = this.getStringParameter('staticTooltipContent');
        
        if (!tooltipContent) {
            tooltipContent = this.getStringParameter('boundTooltipContent');
        }
        
        if (!tooltipContent) {
            tooltipContent = 'No tooltip content provided';
        }
        
        const allowHtml = this.getBooleanParameter('allowHtml', false);
        
        if (allowHtml) {
            this._tooltipElement.innerHTML = this.sanitizeTooltipHtml(tooltipContent);
        } else {
            const preserveLineBreaks = this.getBooleanParameter('preserveLineBreaks', true);
            if (preserveLineBreaks) {
                this._tooltipElement.innerHTML = this.escapeHtml(tooltipContent).replace(/\n/g, '<br>');
            } else {
                this._tooltipElement.textContent = tooltipContent;
            }
        }
    }

    private escapeHtml(text: string): string {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    private sanitizeTooltipHtml(html: string): string {
        if (!html) return '';
        
        const allowedTags = ['b', 'i', 'u', 'strong', 'em', 'br', 'span', 'div', 'p', 
                            'ul', 'ol', 'li', 'img', 'a', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'];
    
        const allowedAttributes: Record<string, string[]> = {
            'img': ['src', 'alt', 'width', 'height', 'style', 'title'],
            'a': ['href', 'target', 'title', 'rel'],
            'span': ['style', 'class'],
            'div': ['style', 'class'],
            'p': ['style', 'class']
        };
        
        // Remove dangerous content
        let sanitized = html
            .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
            .replace(/javascript:/gi, '')
            .replace(/vbscript:/gi, '')
            .replace(/on\w+\s*=\s*["'][^"']*["']/gi, '');
        
        // Basic tag filtering
        sanitized = sanitized.replace(/<\/?(\w+)([^>]*)>/gi, (match, tagName, attributes) => {
            const tag = tagName.toLowerCase();
            
            if (!allowedTags.includes(tag)) {
                return '';
            }
            
            if (match.startsWith('</')) {
                return `</${tag}>`;
            }
            
            return `<${tag}${attributes}>`;
        });
        
        return sanitized;
    }

    private getDefaultIcon(iconType: string): string {
        const icons: Record<string, string> = {
            'info': 'ℹ',
            'question': '?',
            'warning': '⚠',
            'error': '❗',
            'help': '❓'
        };
        
        return icons[iconType.toLowerCase()] || icons['info'];
    }

    // FIXED: Improved showTooltip method
    private showTooltip(): void {
        if (!this._tooltipElement) {
            console.error("Tooltip element not available");
            return;
        }
        
        console.log("Showing tooltip");
        
        // Clear any hide timeout
        this.clearAllTimeouts();
        
        // Set state
        this._isTooltipVisible = true;
        
        // Highlight the icon
        if (this._iconElement) {
            this._iconElement.classList.add('highlighted');
        }
        
        // Update tooltip content before showing (in case it changed)
        this.updateTooltipContent();
        
        // Position and show tooltip
        this.positionTooltip();
        
        // Use RAF for smooth animation
        requestAnimationFrame(() => {
            if (this._tooltipElement && this._isTooltipVisible) {
                this._tooltipElement.classList.add('visible');
                
                // Force background color application
                const bgColor = this.getStringParameter('tooltipBackgroundColor', '#323130');
                const textColor = this.getStringParameter('tooltipTextColor', '#ffffff');
                
                this._tooltipElement.style.setProperty('background-color', bgColor, 'important');
                this._tooltipElement.style.setProperty('color', textColor, 'important');
                this._tooltipElement.style.setProperty('visibility', 'visible', 'important');
                this._tooltipElement.style.setProperty('opacity', '1', 'important');
            }
        });
    }

    // FIXED: Improved hideTooltip method
    private hideTooltip(): void {
        if (!this._tooltipElement) {
            console.error("Tooltip element not available");
            return;
        }
        
        if (!this._isTooltipVisible) {
            console.log("Tooltip already hidden, skipping hide");
            return;
        }
        
        console.log("Hiding tooltip");
        
        // Clear any show timeout
        this.clearAllTimeouts();
        
        // Set state
        this._isTooltipVisible = false;
        
        // Remove icon highlighting
        if (this._iconElement) {
            this._iconElement.classList.remove('highlighted');
        }
        
        // Hide tooltip with animation
        this._tooltipElement.classList.remove('visible');
        
        // Clean up after animation completes
        setTimeout(() => {
            if (this._tooltipElement && !this._isTooltipVisible) {
                this._tooltipElement.style.setProperty('visibility', 'hidden', 'important');
                this._tooltipElement.style.setProperty('opacity', '0', 'important');
            }
        }, 200);
    }

    // FIXED: Enhanced positioning with better overflow handling
    private positionTooltip(): void {
        if (!this._tooltipElement || !this._iconElement) return;

        const iconRect = this._iconElement.getBoundingClientRect();
        const viewport = this.getViewportDimensions();
        
        // Reset transform to get accurate measurements
        this._tooltipElement.style.setProperty('transform', 'none', 'important');
        this._tooltipElement.style.setProperty('visibility', 'visible', 'important');
        this._tooltipElement.style.setProperty('opacity', '0', 'important');
        
        const tooltipRect = this._tooltipElement.getBoundingClientRect();
        const position = this.getStringParameter('position', 'auto').toLowerCase();
        const offset = 12;
        const padding = 10;
        
        let top = 0;
        let left = 0;
        let finalPosition = position;

        // Auto-detect best position if position is 'auto' or if specified position would overflow
        if (position === 'auto' || this.wouldOverflow(position, iconRect, tooltipRect, viewport, offset, padding)) {
            finalPosition = this.getBestPosition(iconRect, tooltipRect, viewport, offset, padding);
        }

        // Calculate position based on final determined position
        switch (finalPosition) {
            case 'bottom':
                top = iconRect.bottom + offset;
                left = iconRect.left + (iconRect.width / 2) - (tooltipRect.width / 2);
                break;
            case 'left':
                top = iconRect.top + (iconRect.height / 2) - (tooltipRect.height / 2);
                left = iconRect.left - tooltipRect.width - offset;
                break;
            case 'right':
                top = iconRect.top + (iconRect.height / 2) - (tooltipRect.height / 2);
                left = iconRect.right + offset;
                break;
            case 'top':
            default:
                top = iconRect.top - tooltipRect.height - offset;
                left = iconRect.left + (iconRect.width / 2) - (tooltipRect.width / 2);
                break;
        }

        // Final viewport boundary adjustments
        if (left < padding) {
            left = padding;
        } else if (left + tooltipRect.width > viewport.width - padding) {
            left = Math.max(padding, viewport.width - tooltipRect.width - padding);
        }

        if (top < padding) {
            top = padding;
        } else if (top + tooltipRect.height > viewport.height - padding) {
            top = Math.max(padding, viewport.height - tooltipRect.height - padding);
        }

        // Apply position
        this._tooltipElement.style.setProperty('top', `${Math.round(top)}px`, 'important');
        this._tooltipElement.style.setProperty('left', `${Math.round(left)}px`, 'important');
        
        // Update arrow position based on final placement
        this.updateArrowPosition(finalPosition, iconRect, { top, left, width: tooltipRect.width, height: tooltipRect.height });
    }

    // FIXED: New method to check if position would cause overflow
    private wouldOverflow(
        position: string, 
        iconRect: DOMRect, 
        tooltipRect: DOMRect, 
        viewport: { width: number; height: number }, 
        offset: number, 
        padding: number
    ): boolean {
        let testTop = 0;
        let testLeft = 0;

        switch (position) {
            case 'bottom':
                testTop = iconRect.bottom + offset;
                testLeft = iconRect.left + (iconRect.width / 2) - (tooltipRect.width / 2);
                break;
            case 'left':
                testTop = iconRect.top + (iconRect.height / 2) - (tooltipRect.height / 2);
                testLeft = iconRect.left - tooltipRect.width - offset;
                break;
            case 'right':
                testTop = iconRect.top + (iconRect.height / 2) - (tooltipRect.height / 2);
                testLeft = iconRect.right + offset;
                break;
            case 'top':
                testTop = iconRect.top - tooltipRect.height - offset;
                testLeft = iconRect.left + (iconRect.width / 2) - (tooltipRect.width / 2);
                break;
            default:
                return false;
        }

        return (
            testLeft < padding ||
            testLeft + tooltipRect.width > viewport.width - padding ||
            testTop < padding ||
            testTop + tooltipRect.height > viewport.height - padding
        );
    }

    // FIXED: New method to find the best position automatically
    private getBestPosition(
        iconRect: DOMRect, 
        tooltipRect: DOMRect, 
        viewport: { width: number; height: number }, 
        offset: number, 
        padding: number
    ): string {
        const positions = ['top', 'bottom', 'right', 'left'];
        
        for (const pos of positions) {
            if (!this.wouldOverflow(pos, iconRect, tooltipRect, viewport, offset, padding)) {
                return pos;
            }
        }
        
        // If all positions would overflow, choose the one with most space
        const spaceTop = iconRect.top;
        const spaceBottom = viewport.height - iconRect.bottom;
        const spaceLeft = iconRect.left;
        const spaceRight = viewport.width - iconRect.right;
        
        const maxSpace = Math.max(spaceTop, spaceBottom, spaceLeft, spaceRight);
        
        if (maxSpace === spaceTop) return 'top';
        if (maxSpace === spaceBottom) return 'bottom';
        if (maxSpace === spaceRight) return 'right';
        return 'left';
    }

    // FIXED: New method to update arrow position based on tooltip placement
    private updateArrowPosition(
        position: string, 
        iconRect: DOMRect, 
        tooltipRect: { top: number; left: number; width: number; height: number }
    ): void {
        if (!this._tooltipElement) return;

        const bgColor = this.getStringParameter('tooltipBackgroundColor', '#323130');
        
        // Remove existing arrow styles
        const existingArrow = this._tooltipElement.querySelector('::before');
        
        // Calculate arrow position based on tooltip placement
        let arrowStyle = '';
        
        switch (position) {
            case 'bottom':
                // Arrow points up
                arrowStyle = `
                    .tooltip-native-content::before {
                        content: '' !important;
                        position: absolute !important;
                        width: 0 !important;
                        height: 0 !important;
                        border: 6px solid transparent !important;
                        border-bottom-color: ${bgColor} !important;
                        top: -12px !important;
                        left: ${Math.max(12, Math.min(tooltipRect.width - 12, iconRect.left + (iconRect.width / 2) - tooltipRect.left))}px !important;
                        transform: none !important;
                    }
                `;
                break;
            case 'top':
                // Arrow points down
                arrowStyle = `
                    .tooltip-native-content::before {
                        content: '' !important;
                        position: absolute !important;
                        width: 0 !important;
                        height: 0 !important;
                        border: 6px solid transparent !important;
                        border-top-color: ${bgColor} !important;
                        bottom: -12px !important;
                        left: ${Math.max(12, Math.min(tooltipRect.width - 12, iconRect.left + (iconRect.width / 2) - tooltipRect.left))}px !important;
                        transform: none !important;
                    }
                `;
                break;
            case 'right':
                // Arrow points left
                arrowStyle = `
                    .tooltip-native-content::before {
                        content: '' !important;
                        position: absolute !important;
                        width: 0 !important;
                        height: 0 !important;
                        border: 6px solid transparent !important;
                        border-right-color: ${bgColor} !important;
                        left: -12px !important;
                        top: ${Math.max(12, Math.min(tooltipRect.height - 12, iconRect.top + (iconRect.height / 2) - tooltipRect.top))}px !important;
                        transform: none !important;
                    }
                `;
                break;
            case 'left':
                // Arrow points right
                arrowStyle = `
                    .tooltip-native-content::before {
                        content: '' !important;
                        position: absolute !important;
                        width: 0 !important;
                        height: 0 !important;
                        border: 6px solid transparent !important;
                        border-left-color: ${bgColor} !important;
                        right: -12px !important;
                        top: ${Math.max(12, Math.min(tooltipRect.height - 12, iconRect.top + (iconRect.height / 2) - tooltipRect.top))}px !important;
                        transform: none !important;
                    }
                `;
                break;
        }
        
        // Update the arrow style in the existing style element
        if (this._styleElement && arrowStyle) {
            // Find and replace the arrow style in the existing styles
            let styleContent = this._styleElement.textContent || '';
            const arrowRegex = /\/\* Tooltip arrow \*\/[\s\S]*?(?=\/\*|$)/;
            
            if (arrowRegex.test(styleContent)) {
                styleContent = styleContent.replace(arrowRegex, `/* Tooltip arrow */\n${arrowStyle}`);
            } else {
                styleContent += `\n/* Tooltip arrow */\n${arrowStyle}`;
            }
            
            this._styleElement.textContent = styleContent;
        }
    }

    private getViewportDimensions(): { width: number; height: number } {
        return {
            width: Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0),
            height: Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0)
        };
    }

    private handleIconClick(event: MouseEvent): void {
        event.preventDefault();
        event.stopPropagation();
        
        const redirectUrl = this.getStringParameter('redirectUrl');
        
        if (redirectUrl && redirectUrl.trim()) {
            const sanitizedUrl = this.sanitizeUrl(redirectUrl.trim());
            
            if (sanitizedUrl) {
                const openInNewTab = this.getBooleanParameter('openInNewTab', true);
                
                try {
                    if (openInNewTab) {
                        window.open(sanitizedUrl, '_blank', 'noopener,noreferrer');
                    } else {
                        window.location.href = sanitizedUrl;
                    }
                } catch (error) {
                    console.error('Error navigating to URL:', error);
                }
            }
        }
        
        // Always hide tooltip after click
        this._isHovering = false;
        this.hideTooltip();
    }

    private sanitizeUrl(url: string): string | null {
        if (!url) return null;
        
        try {
            // Handle relative URLs
            if (url.startsWith('/') || url.startsWith('./') || url.startsWith('../')) {
                return url;
            }
            
            // Handle absolute URLs
            if (url.startsWith('http://') || url.startsWith('https://')) {
                new URL(url); // Validate URL
                return url;
            }
            
            // Handle special protocols
            if (url.startsWith('mailto:') || url.startsWith('tel:')) {
                return url;
            }

            // Add https:// if no protocol specified
            if (!url.includes('://') && !url.startsWith('/')) {
                const urlWithProtocol = 'https://' + url;
                new URL(urlWithProtocol); // Validate
                return urlWithProtocol;
            }
            
            return null;
        } catch (error) {
            console.warn('Invalid URL:', url, error);
            return null;
        }
    }

    // FIXED: Improved resize handling
    private handleResize(): void {
        if (this._isTooltipVisible) {
            // Clear existing timeout
            this.clearAllTimeouts();
            
            // Reposition with small delay to avoid excessive calculations
            this._showTimeout = window.setTimeout(() => {
                if (this._isTooltipVisible) {
                    this.positionTooltip();
                }
                this._showTimeout = null;
            }, 50);
        }
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;

        // Update input field value if it changed externally
        const newValue = this.getStringParameter('dummyField', '');
        if (newValue !== this._currentValue && this._inputElement) {
            this._currentValue = newValue;
            this._inputElement.value = newValue;
        }

        // Update visibility
        const isVisible = true; // Assuming control is always visible unless you have a visibility parameter
        if (this._container) {
            this._container.style.display = isVisible ? 'block' : 'none';
        }

        // Update tooltip content
        this.updateTooltipContent();
        this.updateIconContent();
        this.updateInputProperties();
        this.updateStyles();

        // Re-position tooltip if visible
        if (this._isTooltipVisible) {
            this.positionTooltip();
        }
    }

    private updateInputProperties(): void {
        if (!this._inputElement) return;

        const placeholder = this.getStringParameter('placeholder', '');
        this._inputElement.placeholder = placeholder;
        
        const isReadonly = this.getBooleanParameter('isReadonly', false);
        this._inputElement.readOnly = isReadonly;
        
        const maxLength = this.getNumberParameter('maxLength', 0);
        if (maxLength > 0) {
            this._inputElement.maxLength = maxLength;
        } else {
            this._inputElement.removeAttribute('maxlength');
        }
    }

    private updateStyles(): void {
        const bgColor = this.getStringParameter('tooltipBackgroundColor', '#323130');
        const textColor = this.getStringParameter('tooltipTextColor', '#ffffff');
        const iconColor = this.getStringParameter('iconColor', '#605e5c');
        const iconSize = this.getNumberParameter('iconSize', 16);
        
        // Update icon styling
        if (this._iconElement) {
            this._iconElement.style.setProperty('color', iconColor, 'important');
            this._iconElement.style.setProperty('font-size', `${iconSize}px`, 'important');
        }

        // Update tooltip styling
        if (this._tooltipElement) {
            this._tooltipElement.style.setProperty('background-color', bgColor, 'important');
            this._tooltipElement.style.setProperty('color', textColor, 'important');
        }
    }

    public getOutputs(): IOutputs {
        return {
            dummyField: this._currentValue
        };
    }

    public destroy(): void {
        console.log("Destroying tooltip control");
        
        // Clean up state
        this._isHovering = false;
        this._isTooltipVisible = false;
        
        // Clean up timeouts
        this.clearAllTimeouts();

        // Clean up ResizeObserver
        if (this._observer) {
            this._observer.disconnect();
            this._observer = null;
        }

        // Remove event listeners
        if (this._iconElement) {
            this._iconElement.removeEventListener('mouseenter', this.handleIconMouseEnter.bind(this));
            this._iconElement.removeEventListener('mouseleave', this.handleIconMouseLeave.bind(this));
            this._iconElement.removeEventListener('click', this._clickHandlerBound);
            this._iconElement.removeEventListener('focus', this._showTooltipBound);
            this._iconElement.removeEventListener('blur', this._hideTooltipBound);
        }

        if (this._inputElement) {
            this._inputElement.removeEventListener('input', this._inputChangeBound);
            this._inputElement.removeEventListener('change', this._inputChangeBound);
            this._inputElement.removeEventListener('blur', this._inputChangeBound);
        }

        document.removeEventListener('click', this._documentClickBound);
        window.removeEventListener('resize', this._resizeHandlerBound);
        window.removeEventListener('scroll', this._resizeHandlerBound, true);
        
        // Remove DOM elements
        if (this._styleElement && this._styleElement.parentNode) {
            this._styleElement.parentNode.removeChild(this._styleElement);
        }

        if (this._tooltipElement && this._tooltipElement.parentNode) {
            this._tooltipElement.parentNode.removeChild(this._tooltipElement);
        }

        // Clean up references
        this._container = null!;
        this._context = null!;
        this._tooltipElement = null!;
        this._iconElement = null!;
        this._inputElement = null!;
        this._inputContainer = null!;
        this._styleElement = null!;
        this._notifyOutputChanged = null!;
    }
}