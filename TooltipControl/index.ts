import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class TooltipControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _tooltipElement: HTMLDivElement;
    private _iconElement: HTMLElement;
    private _styleElement: HTMLStyleElement;
    private _showTooltipBound: () => void;
    private _hideTooltipBound: () => void;
    private _clickHandlerBound: (event: MouseEvent) => void;
    private _resizeHandlerBound: () => void;
    private _hideTimeout: number | null = null;
    private _lastBgColor = '';
    private _lastTextColor = '';
    private _positionUpdateFrameId: number | null = null;
    private _hasRedirectUrl = false;

    constructor() {
        this._showTooltipBound = this.showTooltip.bind(this);
        this._hideTooltipBound = this.hideTooltip.bind(this);
        this._clickHandlerBound = this.handleIconClick.bind(this);
        this._resizeHandlerBound = this.handleResize.bind(this);
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

        this.createTooltipControl();

        this.addWindowResizeHandler();
    }

    private injectGlobalStyles(): void {
        if (this._styleElement && this._styleElement.parentNode) {
            this._styleElement.parentNode.removeChild(this._styleElement);
        }

        this._styleElement = document.createElement('style');
        this._styleElement.id = `tooltip-styles-${Date.now()}`;
        
        const bgColor = this.getStringParameter('tooltipBackgroundColor', '#333333');
        const textColor = this.getStringParameter('tooltipTextColor', '#ffffff');
        const linkColor = this.getLinkColor(textColor);
        
        document.documentElement.style.setProperty('--tooltip-bg-color', bgColor);
        document.documentElement.style.setProperty('--tooltip-text-color', textColor);
        document.documentElement.style.setProperty('--tooltip-link-color', linkColor);
        
        this._lastBgColor = bgColor;
        this._lastTextColor = textColor;
        
        this._styleElement.textContent = this.generateEnhancedStyleContent(bgColor, textColor, linkColor);
        
        document.head.appendChild(this._styleElement);
    }

    private generateEnhancedStyleContent(bgColor: string, textColor: string, linkColor: string): string {
        return `
            .tooltip-content,
            body .tooltip-content,
            html .tooltip-content,
            div.tooltip-content {
                position: fixed !important;
                background-color: ${bgColor} !important;
                background: ${bgColor} !important;
                background-image: none !important;
                color: ${textColor} !important;
                padding: 12px 16px !important;
                border-radius: 4px !important;
                font-size: 13px !important;
                line-height: 1.5 !important;
                max-width: 400px !important;
                min-width: 200px !important;
                width: auto !important;
                height: auto !important;
                z-index: 2147483647 !important;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3) !important;
                visibility: hidden !important;
                opacity: 0 !important;
                pointer-events: none !important;
                word-wrap: break-word !important;
                border: 1px solid rgba(255, 255, 255, 0.1) !important;
                white-space: normal !important;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif !important;
                filter: none !important;
                backdrop-filter: none !important;
                -webkit-backdrop-filter: none !important;
                background-blend-mode: normal !important;
                mix-blend-mode: normal !important;
                transition: opacity 0.2s ease, visibility 0.2s ease !important;
                overflow-wrap: break-word !important;
                word-break: break-word !important;
                box-sizing: border-box !important;
                margin: 0 !important;
                text-align: left !important;
                text-transform: none !important;
                letter-spacing: normal !important;
                text-shadow: none !important;
                
                /* Better positioning stability */
                transform: translateZ(0) !important;
                -webkit-transform: translateZ(0) !important;
                backface-visibility: hidden !important;
                -webkit-backface-visibility: hidden !important;
                will-change: opacity, visibility !important;
                contain: layout style paint !important;
                
                /* Force GPU acceleration for smooth positioning */
                -webkit-font-smoothing: antialiased !important;
                -moz-osx-font-smoothing: grayscale !important;
                text-rendering: optimizeLegibility !important;
                
                /* Prevent layout shifts */
                overflow: hidden !important;
                overscroll-behavior: none !important;
            }
            
            .tooltip-content.visible,
            body .tooltip-content.visible,
            html .tooltip-content.visible,
            div.tooltip-content.visible {
                visibility: visible !important;
                opacity: 1 !important;
                background-color: ${bgColor} !important;
                background: ${bgColor} !important;
                background-image: none !important;
                display: block !important;
                
                /* Ensure visibility override */
                -webkit-transform: translateZ(0) !important;
                transform: translateZ(0) !important;
            }
            
            .tooltip-wrapper {
                position: relative !important;
                display: inline-block !important;
                vertical-align: middle !important;
            }
            
            .tooltip-icon {
                cursor: pointer !important;
                user-select: none !important;
                -webkit-user-select: none !important;
                -moz-user-select: none !important;
                -ms-user-select: none !important;
                display: inline-block !important;
                margin-left: 5px !important;
                vertical-align: middle !important;
                transition: color 0.2s ease, background-color 0.2s ease, transform 0.1s ease !important;
                outline: none !important;
                border-radius: 50% !important;
                padding: 2px !important;
                font-weight: bold !important;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif !important;
                line-height: 1 !important;
                box-sizing: border-box !important;
                pointer-events: auto !important;
                position: relative !important;
                z-index: 1 !important;
            }
            
            .tooltip-icon:hover,
            .tooltip-icon:focus {
                color: #007acc !important;
                background-color: rgba(0, 122, 204, 0.1) !important;
                transform: scale(1.1) !important;
            }
            
            .tooltip-icon:active {
                transform: scale(0.95) !important;
            }
            
            .tooltip-icon.clickable {
                cursor: pointer !important;
                border: 1px solid transparent !important;
            }
            
            .tooltip-icon.clickable:hover {
                border-color: rgba(0, 122, 204, 0.3) !important;
                box-shadow: 0 2px 8px rgba(0, 122, 204, 0.2) !important;
            }
            
            /* Enhanced HTML content styling */
            .tooltip-content h1, .tooltip-content h2, .tooltip-content h3, 
            .tooltip-content h4, .tooltip-content h5, .tooltip-content h6 {
                margin: 8px 0 4px 0 !important;
                padding: 0 !important;
                font-weight: bold !important;
                line-height: 1.3 !important;
                color: ${textColor} !important;
                background: none !important;
                border: none !important;
                text-align: left !important;
            }
            
            .tooltip-content h1 { font-size: 16px !important; }
            .tooltip-content h2 { font-size: 15px !important; }
            .tooltip-content h3 { font-size: 14px !important; }
            .tooltip-content h4, .tooltip-content h5, .tooltip-content h6 { font-size: 13px !important; }
            
            .tooltip-content p {
                margin: 6px 0 !important;
                padding: 0 !important;
                line-height: 1.5 !important;
                color: ${textColor} !important;
                background: none !important;
                border: none !important;
                text-align: left !important;
            }
            
            .tooltip-content ul, .tooltip-content ol {
                margin: 8px 0 !important;
                padding-left: 20px !important;
                line-height: 1.4 !important;
                color: ${textColor} !important;
                background: none !important;
                border: none !important;
            }
            
            .tooltip-content li {
                margin: 2px 0 !important;
                padding: 0 !important;
                color: ${textColor} !important;
                background: none !important;
                border: none !important;
            }
            
            .tooltip-content ul { list-style-type: disc !important; }
            .tooltip-content ol { list-style-type: decimal !important; }
            
            .tooltip-content ul ul, .tooltip-content ol ol,
            .tooltip-content ul ol, .tooltip-content ol ul {
                margin: 2px 0 !important;
                padding-left: 16px !important;
            }
            
            .tooltip-content img {
                max-width: 100% !important;
                height: auto !important;
                border-radius: 3px !important;
                margin: 4px 0 !important;
                display: block !important;
                border: none !important;
                background: none !important;
            }
            
            .tooltip-content a {
                color: ${linkColor} !important;
                text-decoration: underline !important;
                background: none !important;
                border: none !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            
            .tooltip-content a:hover {
                opacity: 0.8 !important;
                background: none !important;
            }
            
            .tooltip-content strong, .tooltip-content b {
                font-weight: bold !important;
                color: ${textColor} !important;
                background: none !important;
                border: none !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            
            .tooltip-content em, .tooltip-content i {
                font-style: italic !important;
                color: ${textColor} !important;
                background: none !important;
                border: none !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            
            .tooltip-content u {
                text-decoration: underline !important;
                color: ${textColor} !important;
                background: none !important;
                border: none !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            
            .tooltip-content > *:first-child { margin-top: 0 !important; }
            .tooltip-content > *:last-child { margin-bottom: 0 !important; }

            /* Better responsive behavior */
            @media screen and (max-width: 480px) {
                .tooltip-content {
                    max-width: calc(100vw - 40px) !important;
                    min-width: 150px !important;
                    font-size: 12px !important;
                    padding: 10px 12px !important;
                }
            }
            
            @media screen and (max-width: 320px) {
                .tooltip-content {
                    max-width: calc(100vw - 20px) !important;
                    min-width: 120px !important;
                    font-size: 11px !important;
                    padding: 8px 10px !important;
                }
            }
        `;
    }

    private getLinkColor(textColor: string): string {
        
        if (textColor.toLowerCase() === '#ffffff' || textColor.toLowerCase() === 'white') {
            return '#87ceeb'; 
        } else {
            return '#0066cc'; 
        }
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

    private createTooltipControl(): void {
        
        if (!this._container) {
            console.error("Container is not available");
            return;
        }

        this._container.innerHTML = '';

        const wrapper = document.createElement('div');
        wrapper.className = 'tooltip-wrapper';

        this._iconElement = document.createElement('span');
        this._iconElement.className = 'tooltip-icon';
        
        const iconColor = this.getStringParameter('iconColor', '#666666');
        const iconSize = this.getNumberParameter('iconSize', 16);
        
        this._iconElement.style.cssText = `
            cursor: pointer;
            font-size: ${iconSize}px;
            color: ${iconColor};
            line-height: 1;
            user-select: none;
            display: inline-block;
        `;
        
        this.updateIconContent();

        this.updateRedirectFunctionality();

        this._tooltipElement = document.createElement('div');
        this._tooltipElement.className = 'tooltip-content';
        
        const bgColor = this.getStringParameter('tooltipBackgroundColor', '#333333');
        const textColor = this.getStringParameter('tooltipTextColor', '#ffffff');
        
        this._tooltipElement.style.setProperty('background-color', bgColor, 'important');
        this._tooltipElement.style.setProperty('background', bgColor, 'important');
        this._tooltipElement.style.setProperty('background-image', 'none', 'important');
        this._tooltipElement.style.setProperty('color', textColor, 'important');
        this._tooltipElement.style.setProperty('filter', 'none', 'important');
        this._tooltipElement.style.setProperty('backdrop-filter', 'none', 'important');

        this.updateTooltipContent();

        this._iconElement.addEventListener('mouseenter', this._showTooltipBound);
        this._iconElement.addEventListener('mouseleave', this._hideTooltipBound);
        this._iconElement.addEventListener('mouseover', this._showTooltipBound);
        this._iconElement.addEventListener('mouseout', this._hideTooltipBound);
        this._iconElement.addEventListener('focus', this._showTooltipBound);
        this._iconElement.addEventListener('blur', this._hideTooltipBound);
        this._iconElement.addEventListener('click', this._clickHandlerBound);

        this._iconElement.setAttribute('tabindex', '0');
        this._iconElement.setAttribute('role', 'button');
        this._iconElement.setAttribute('aria-describedby', 'tooltip-content');
        
        this.updateAccessibilityAttributes();

        wrapper.appendChild(this._iconElement);
        wrapper.appendChild(this._tooltipElement);
        this._container.appendChild(wrapper);
    }

    private updateRedirectFunctionality(): void {
        const redirectUrl = this.getStringParameter('redirectUrl');
        this._hasRedirectUrl = Boolean(redirectUrl && redirectUrl.trim());
        
        if (this._iconElement) {
            if (this._hasRedirectUrl) {
                this._iconElement.classList.add('clickable');
                this._iconElement.style.cursor = 'pointer';
            } else {
                this._iconElement.classList.remove('clickable');
                this._iconElement.style.cursor = 'help';
            }
        }
    }

    private updateAccessibilityAttributes(): void {
        if (!this._iconElement) return;
        
        const redirectUrl = this.getStringParameter('redirectUrl');
        const hasRedirect = Boolean(redirectUrl && redirectUrl.trim());
        
        if (hasRedirect) {
            this._iconElement.setAttribute('aria-label', 'Show tooltip and click to navigate');
        } else {
            this._iconElement.setAttribute('aria-label', 'Show tooltip');
        }
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
                    const newWindow = window.open(sanitizedUrl, '_blank', 'noopener,noreferrer');
                    if (!newWindow) {
                        console.warn('Popup blocked by browser');
                        alert('Popup blocked. Please allow popups for this site or manually navigate to: ' + sanitizedUrl);
                        return;
                    }
                } else {
                    window.location.href = sanitizedUrl;
                }
            } catch (error) {
                console.error('Error navigating to URL:', error);
                if (openInNewTab) {
                    alert('Unable to open link in new tab. Please check your browser settings or manually navigate to: ' + sanitizedUrl);
                } else {
                    alert('Unable to navigate to the specified link. Please check the URL and try again.');
                }
            }
        } else {
            console.error('Invalid URL provided:', redirectUrl);
            alert('Invalid URL provided. Please check the redirect URL configuration.');
        }
    }
    
    this.hideTooltip();
}

    private sanitizeUrl(url: string): string | null {
        if (!url) return null;
        
        try {
            if (url.startsWith('/') || url.startsWith('./') || url.startsWith('../')) {
                return url;
            }
            
            if (url.startsWith('http://') || url.startsWith('https://')) {
                const urlObj = new URL(url);
                return urlObj.href;
            }
            
            if (url.startsWith('mailto:') || url.startsWith('tel:')) {
                return url;
            }

            if (!url.includes('://') && !url.startsWith('/')) {
                try {
                    const urlWithProtocol = 'https://' + url;
                    const urlObj = new URL(urlWithProtocol);
                    return urlObj.href;
                } catch {
                    return null;
                }
            }
            
            return null;
        } catch (error) {
            console.warn('Error sanitizing URL:', error);
            return null;
        }
    }

    private updateIconContent(): void {
        if (!this._iconElement) return;

        const iconType = this.getStringParameter('iconType', 'info');
        
        this._iconElement.textContent = this.getDefaultIcon(iconType);
    }

    private updateTooltipContent(): void {
        if (!this._tooltipElement) return;

        let tooltipContent = this.getStringParameter('staticTooltipContent');
        
        if (!tooltipContent) {
            tooltipContent = this.getStringParameter('boundTooltipContent');
        }
        
        if (!tooltipContent) {
            tooltipContent = 'Tooltip content';
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
            'p': ['style', 'class'],
            'ul': ['style', 'class', 'type'],
            'ol': ['style', 'class', 'type', 'start'],
            'li': ['style', 'class', 'value'],
            'h1': ['style', 'class'],
            'h2': ['style', 'class'],
            'h3': ['style', 'class'],
            'h4': ['style', 'class'],
            'h5': ['style', 'class'],
            'h6': ['style', 'class'],
            'strong': ['style', 'class'],
            'b': ['style', 'class'],
            'i': ['style', 'class'],
            'em': ['style', 'class'],
            'u': ['style', 'class']
        };
        
        let sanitized = html
            .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
            .replace(/javascript:/gi, '')
            .replace(/vbscript:/gi, '')
            .replace(/on\w+\s*=\s*["'][^"']*["']/gi, '');
        
        sanitized = sanitized.replace(/data:(?!image\/(png|jpg|jpeg|gif|svg\+xml);base64,)[^"']*/gi, '');
        
        sanitized = sanitized.replace(/<\/?(\w+)([^>]*)>/gi, (match, tagName, attributes) => {
            const tag = tagName.toLowerCase();
            
            if (!allowedTags.includes(tag)) {
                return ''; 
            }
            
            if (match.startsWith('</')) {
                return `</${tag}>`;
            }
            
            if (attributes && allowedAttributes[tag]) {
                const allowedAttrs = allowedAttributes[tag];
                const processedAttributes = attributes.replace(/(\w+)\s*=\s*["']([^"']*)["']/gi, 
                    (attrMatch: string, attrName: string, attrValue: string) => {
                        const attr = attrName.toLowerCase();
                        
                        if (!allowedAttrs.includes(attr)) {
                            return ''; 
                        }
                        
                        if (attr === 'href' || attr === 'src') {
                            if (attrValue.match(/^(https?:\/\/|\/|#|\?)/i) || 
                                (attr === 'src' && attrValue.match(/^data:image\/(png|jpg|jpeg|gif|svg\+xml);base64,/i))) {
                                return ` ${attr}="${attrValue}"`;
                            }
                            return '';
                        }
                        
                        if (attr === 'target' && attrValue !== '_blank') {
                            return ' target="_blank"';
                        }
                        
                        if (attr === 'rel' && attrValue.indexOf('noopener') === -1) {
                            return ' rel="noopener noreferrer"'; 
                        }
                        
                        if (attr === 'style') {
                            const safeStyle = this.sanitizeStyle(attrValue);
                            return safeStyle ? ` style="${safeStyle}"` : '';
                        }
                        
                        return ` ${attr}="${attrValue}"`;
                    });
                
                return `<${tag}${processedAttributes}>`;
            } else if (allowedTags.includes(tag)) {
                return `<${tag}>`;
            }
            
            return '';
        });
        
        return sanitized;
    }

    private sanitizeStyle(style: string): string {
        if (!style) return '';
        
        const dangerousProps = ['expression', 'javascript', 'vbscript', 'import', 'behavior'];
        let sanitizedStyle = style;
        
        dangerousProps.forEach(prop => {
            const regex = new RegExp(`${prop}\\s*\\([^)]*\\)`, 'gi');
            sanitizedStyle = sanitizedStyle.replace(regex, '');
        });
        
        sanitizedStyle = sanitizedStyle.replace(/url\s*\(\s*(?!['"]?data:image\/)[^)]*\)/gi, '');
        
        return sanitizedStyle;
    }

    private getDefaultIcon(iconType: string): string {
        const icons: Record<string, string> = {
            'info': 'ℹ',
            'question': '?',
            'warning': '⚠',
            'error': '!'
        };
        
        return icons[iconType.toLowerCase()] || icons['info'];
    }

    private showTooltip(): void {
        if (!this._tooltipElement) return;
        
        console.log("Showing tooltip with enhanced timing");
        
        if (this._hideTimeout) {
            clearTimeout(this._hideTimeout);
            this._hideTimeout = null;
        }
        
        if (this._positionUpdateFrameId) {
            cancelAnimationFrame(this._positionUpdateFrameId);
            this._positionUpdateFrameId = null;
        }
        
        this._positionUpdateFrameId = requestAnimationFrame(() => {
            if (!this._tooltipElement) return;
            
            this.positionTooltip();
            
            this._positionUpdateFrameId = requestAnimationFrame(() => {
                if (!this._tooltipElement) return;
                
                this._tooltipElement.classList.add('visible');
                
                const bgColor = this.getStringParameter('tooltipBackgroundColor', '#333333');
                const textColor = this.getStringParameter('tooltipTextColor', '#ffffff');
                
                const forceStyles = {
                    backgroundColor: bgColor,
                    background: bgColor,
                    backgroundImage: 'none',
                    color: textColor,
                    visibility: 'visible',
                    opacity: '1',
                    display: 'block',
                    filter: 'none',
                    backdropFilter: 'none'
                };
                
                Object.entries(forceStyles).forEach(([prop, value]) => {
                    this._tooltipElement!.style.setProperty(prop, value, 'important');
                });
            });
        });
    }

    private hideTooltip(): void {
        if (!this._tooltipElement) return;
        
        console.log("Hiding tooltip with enhanced timing");
        
        if (this._positionUpdateFrameId) {
            cancelAnimationFrame(this._positionUpdateFrameId);
            this._positionUpdateFrameId = null;
        }
        
        this._hideTimeout = window.setTimeout(() => {
            if (!this._tooltipElement) return;
            
            this._tooltipElement.classList.remove('visible');
            this._tooltipElement.style.setProperty('visibility', 'hidden', 'important');
            this._tooltipElement.style.setProperty('opacity', '0', 'important');
            this._hideTimeout = null;
        }, 100);
    }

    private positionTooltip(): void {
        if (!this._tooltipElement || !this._iconElement) return;

        requestAnimationFrame(() => {
            this.performPositioning();
        });
    }

    private performPositioning(): void {
        if (!this._tooltipElement || !this._iconElement) return;

        try {
            if (!document.contains(this._tooltipElement)) {
                document.body.appendChild(this._tooltipElement);
            }

            this._tooltipElement.style.position = 'fixed';
            this._tooltipElement.style.top = '-9999px';
            this._tooltipElement.style.left = '-9999px';
            this._tooltipElement.style.visibility = 'hidden';
            this._tooltipElement.style.opacity = '0';
            this._tooltipElement.style.display = 'block';
            this._tooltipElement.style.maxWidth = '400px';
            this._tooltipElement.style.width = 'auto';
            this._tooltipElement.style.transform = 'none';
            this._tooltipElement.style.zIndex = '2147483647';

            void this._tooltipElement.offsetHeight;
            void this._tooltipElement.offsetWidth;
            
            requestAnimationFrame(() => {
                this.calculateAndApplyPosition();
            });
            
        } catch (error) {
            console.error('Error positioning tooltip:', error);
            this.applyFallbackPosition();
        }
    }

    private getStableIconRect(): DOMRect | null {
        if (!this._iconElement) return null;
        
        let rect = this._iconElement.getBoundingClientRect();
        let attempts = 0;
        const maxAttempts = 3;
        
        while (attempts < maxAttempts && (rect.width === 0 || rect.height === 0)) {
            void this._iconElement.offsetHeight;
            rect = this._iconElement.getBoundingClientRect();
            attempts++;
        }
        
        if (rect.width === 0 || rect.height === 0) {
            console.warn('Icon has zero dimensions, using fallback');
            return null;
        }
        
        const adjustedRect = this.adjustRectForPowerPlatform(rect);
        return adjustedRect;
    }

    private adjustRectForPowerPlatform(rect: DOMRect): DOMRect {
        const isInIframe = window !== window.parent;
        const adjustedTop = rect.top;
        const adjustedLeft = rect.left;
        
        console.log(`Original rect: top=${rect.top}, left=${rect.left}, isInIframe=${isInIframe}`);
        
        if (isInIframe) {
            try {
                const iframe = window.frameElement as HTMLIFrameElement;
                if (iframe) {
                    const iframeRect = iframe.getBoundingClientRect();
                    console.log(`Iframe offset: top=${iframeRect.top}, left=${iframeRect.left}`);
                    
                }
            } catch (e) {
                console.log('Cannot access iframe element (cross-origin)');
            }
        }
        
        const formHeader = document.querySelector('.ms-DetailHeader, [data-id="form-header"], .form-header');
        const navigation = document.querySelector('.ms-Nav, [data-id="navbar"], .navbar');
        
        let headerHeight = 0;
        if (formHeader) {
            const headerRect = formHeader.getBoundingClientRect();
            headerHeight = headerRect.height;
            console.log(`Form header height: ${headerHeight}`);
        }
        
        if (navigation) {
            const navRect = navigation.getBoundingClientRect();
            headerHeight += navRect.height;
            console.log(`Navigation height: ${navRect.height}, total header height: ${headerHeight}`);
        }
        
        let scrollableParent = this._iconElement.parentElement;
        let scrollOffset = 0;
        
        while (scrollableParent && scrollableParent !== document.body) {
            const computedStyle = window.getComputedStyle(scrollableParent);
            if (computedStyle.overflow === 'auto' || computedStyle.overflow === 'scroll' || 
                computedStyle.overflowY === 'auto' || computedStyle.overflowY === 'scroll') {
                scrollOffset += scrollableParent.scrollTop;
                console.log(`Found scrollable parent with scrollTop: ${scrollableParent.scrollTop}`);
                break;
            }
            scrollableParent = scrollableParent.parentElement;
        }
        
        console.log(`Adjustments: headerHeight=${headerHeight}, scrollOffset=${scrollOffset}`);
        
        const adjustedRect = {
            top: adjustedTop,
            left: adjustedLeft,
            right: rect.right,
            bottom: rect.bottom,
            width: rect.width,
            height: rect.height,
            x: rect.x,
            y: rect.y,
            toJSON: rect.toJSON
        } as DOMRect;
        
        console.log(`Adjusted rect: top=${adjustedRect.top}, left=${adjustedRect.left}`);
        return adjustedRect;
    }

    private getViewportDimensions(): { width: number; height: number } {
        let width: number;
        let height: number;
        
        const isInIframe = window !== window.parent;
        
        if (isInIframe) {
            width = window.innerWidth || document.documentElement.clientWidth;
            height = window.innerHeight || document.documentElement.clientHeight;
            console.log(`Iframe viewport: ${width}x${height}`);
        } else {
            width = Math.max(
                document.documentElement.clientWidth || 0,
                window.innerWidth || 0,
                document.body.clientWidth || 0
            );
            
            height = Math.max(
                document.documentElement.clientHeight || 0,
                window.innerHeight || 0,
                document.body.clientHeight || 0
            );
            console.log(`Main viewport: ${width}x${height}`);
        }
        
        return { width, height };
    }

    private calculateAndApplyPosition(): void {
        if (!this._tooltipElement || !this._iconElement) return;

        const iconRect = this.getStableIconRect();
        if (!iconRect) {
            this.applyFallbackPosition();
            return;
        }
        
        const tooltipRect = this.getStableTooltipRect();
        if (!tooltipRect) {
            this.applyFallbackPosition();
            return;
        }
        
        const viewport = this.getViewportDimensions();
        const position = this.getStringParameter('position', 'right').toLowerCase();
        const offset = 8; 
        
        let finalTop = 0;
        let finalLeft = 0;
        let actualPosition = position;

        console.log(`=== POWER PLATFORM POSITIONING DEBUG ===`);
        console.log(`Position preference: ${position}`);
        console.log(`Icon rect: x:${iconRect.left.toFixed(1)} y:${iconRect.top.toFixed(1)} w:${iconRect.width.toFixed(1)} h:${iconRect.height.toFixed(1)}`);
        console.log(`Tooltip rect: w:${tooltipRect.width.toFixed(1)} h:${tooltipRect.height.toFixed(1)}`);
        console.log(`Viewport: ${viewport.width}x${viewport.height}`);
        console.log(`Window location: ${window.location.href}`);
        console.log(`Is iframe: ${window !== window.parent}`);
        
        switch (position) {
            case 'bottom':
                finalTop = iconRect.bottom + offset;
                finalLeft = iconRect.left + (iconRect.width / 2) - (tooltipRect.width / 2);
                break;
            case 'left':
                finalTop = iconRect.top + (iconRect.height / 2) - (tooltipRect.height / 2);
                finalLeft = iconRect.left - tooltipRect.width - offset;
                break;
            case 'right':
                finalTop = iconRect.top + (iconRect.height / 2) - (tooltipRect.height / 2);
                finalLeft = iconRect.right + offset;
                break;
            case 'top':
                finalTop = iconRect.top - tooltipRect.height - offset;
                finalLeft = iconRect.left + (iconRect.width / 2) - (tooltipRect.width / 2);
                break;
            default: {
                const positioning = this.calculateSmartPositionForPowerPlatform(iconRect, tooltipRect, viewport, offset);
                finalTop = positioning.top;
                finalLeft = positioning.left;
                actualPosition = positioning.position;
                break;
            }
        }

        console.log(`Initial calculated: x:${finalLeft.toFixed(1)} y:${finalTop.toFixed(1)} (${actualPosition})`);

        const adjustedPosition = this.adjustForBoundariesPowerPlatform(
            { top: finalTop, left: finalLeft, position: actualPosition },
            iconRect,
            tooltipRect,
            viewport,
            offset
        );

        finalTop = adjustedPosition.top;
        finalLeft = adjustedPosition.left;
        actualPosition = adjustedPosition.position;

        this.applyFinalPositionPowerPlatform(finalTop, finalLeft, viewport.width);

        console.log(`Final position: x:${Math.round(finalLeft)} y:${Math.round(finalTop)} (${actualPosition})`);
        
        setTimeout(() => this.verifyPositionPowerPlatform(finalTop, finalLeft), 50);
        
        console.log(`=== END POWER PLATFORM POSITIONING ===`);
    }

    private calculateSmartPositionForPowerPlatform(
        iconRect: DOMRect, 
        tooltipRect: DOMRect, 
        viewport: { width: number; height: number }, 
        offset: number
    ): { top: number; left: number; position: string } {
        
        const padding = 10; 
        const positions = [
            {
                name: 'right',
                top: iconRect.top + (iconRect.height / 2) - (tooltipRect.height / 2),
                left: iconRect.right + offset,
                fits: iconRect.right + offset + tooltipRect.width <= viewport.width - padding
            },
            {
                name: 'bottom',
                top: iconRect.bottom + offset,
                left: iconRect.left + (iconRect.width / 2) - (tooltipRect.width / 2),
                fits: iconRect.bottom + offset + tooltipRect.height <= viewport.height - padding
            },
            {
                name: 'left',
                top: iconRect.top + (iconRect.height / 2) - (tooltipRect.height / 2),
                left: iconRect.left - tooltipRect.width - offset,
                fits: iconRect.left - offset - tooltipRect.width >= padding
            },
            {
                name: 'top',
                top: iconRect.top - tooltipRect.height - offset,
                left: iconRect.left + (iconRect.width / 2) - (tooltipRect.width / 2),
                fits: iconRect.top - offset - tooltipRect.height >= padding
            }
        ];
        
        const bestFit = positions.find(pos => pos.fits);
        
        if (bestFit) {
            return {
                top: bestFit.top,
                left: bestFit.left,
                position: bestFit.name
            };
        }
        
        return {
            top: positions[0].top,
            left: positions[0].left,
            position: positions[0].name
        };
    }

    private adjustForBoundariesPowerPlatform(
        current: { top: number; left: number; position: string },
        iconRect: DOMRect,
        tooltipRect: DOMRect,
        viewport: { width: number; height: number },
        offset: number
    ): { top: number; left: number; position: string } {
        
        const padding = 10;
        let { top, left } = current;
        const { position } = current;

        if (left < padding) {
            left = padding;
        } else if (left + tooltipRect.width > viewport.width - padding) {
            left = viewport.width - tooltipRect.width - padding;
        }

        if (top < padding) {
            top = padding;
        } else if (top + tooltipRect.height > viewport.height - padding) {
            top = viewport.height - tooltipRect.height - padding;
        }

        top = Math.max(padding, Math.min(top, viewport.height - tooltipRect.height - padding));
        left = Math.max(padding, Math.min(left, viewport.width - tooltipRect.width - padding));

        return { top, left, position: position + '-adjusted' };
    }


private applyFinalPositionPowerPlatform(top: number, left: number, viewportWidth: number): void {
    if (!this._tooltipElement) return;

    const finalTop = Math.round(top);
    const finalLeft = Math.round(left);
    
    console.log(`Applying final position: top=${finalTop}px, left=${finalLeft}px`);
    
    this._tooltipElement.style.removeProperty('transform');
    this._tooltipElement.style.removeProperty('translate');
    this._tooltipElement.style.removeProperty('translate3d');
    this._tooltipElement.style.removeProperty('translateX');
    this._tooltipElement.style.removeProperty('translateY');
    this._tooltipElement.style.removeProperty('margin');
    this._tooltipElement.style.removeProperty('margin-top');
    this._tooltipElement.style.removeProperty('margin-left');
    
    const customMaxWidth = this.getNumberParameter('tooltipMaxWidth', 0);
    const customMaxHeight = this.getNumberParameter('tooltipMaxHeight', 0);
    
    let maxWidth = Math.min(400, viewportWidth - 30);
    if (customMaxWidth > 0) {
        maxWidth = Math.min(customMaxWidth, viewportWidth - 30);
    }
    
    const styles: Record<string, string> = {
        position: 'fixed',
        top: `${finalTop}px`,
        left: `${finalLeft}px`,
        right: 'auto',
        bottom: 'auto',
        transform: 'none',
        translate: 'none',
        zIndex: '2147483647',
        maxWidth: `${maxWidth}px`,
        willChange: 'opacity, visibility',
        backfaceVisibility: 'hidden',
        margin: '0',
        marginTop: '0',
        marginLeft: '0',
        marginRight: '0',
        marginBottom: '0',
        display: 'block',
        float: 'none',
        clear: 'none',
        alignSelf: 'auto',
        justifySelf: 'auto',
        gridColumn: 'auto',
        gridRow: 'auto',
        flexShrink: '0',
        flexGrow: '0',
        flexBasis: 'auto'
    };

    if (customMaxHeight > 0) {
        styles.maxHeight = `${customMaxHeight}px`;
        styles.overflowY = 'auto';
    }

    Object.entries(styles).forEach(([prop, value]) => {
        this._tooltipElement!.style.setProperty(prop, value, 'important');
    });
    
    void this._tooltipElement.offsetHeight;
    void this._tooltipElement.offsetWidth;
    void this._tooltipElement.offsetTop;
    void this._tooltipElement.offsetLeft;
    
    requestAnimationFrame(() => {
        if (!this._tooltipElement) return;
        
        const immediateRect = this._tooltipElement.getBoundingClientRect();
        const actualTop = Math.round(immediateRect.top);
        const actualLeft = Math.round(immediateRect.left);
        
        if (Math.abs(actualTop - finalTop) > 5 || Math.abs(actualLeft - finalLeft) > 5) {
            console.warn(`Position drift detected. Forcing correction: expected(${finalLeft}, ${finalTop}) vs actual(${actualLeft}, ${actualTop})`);
            
            this._tooltipElement.style.cssText += `; position: fixed !important; top: ${finalTop}px !important; left: ${finalLeft}px !important; transform: none !important; margin: 0 !important;`;
            
            void this._tooltipElement.offsetHeight;
        }
    });
}

    private verifyPositionPowerPlatform(expectedTop: number, expectedLeft: number): void {
        if (!this._tooltipElement) return;
        
        const actualRect = this._tooltipElement.getBoundingClientRect();
        const actualTop = actualRect.top;
        const actualLeft = actualRect.left;
        
        const topDiff = Math.abs(actualTop - expectedTop);
        const leftDiff = Math.abs(actualLeft - expectedLeft);
        
        console.log(`Position verification:`);
        console.log(`Expected: x:${Math.round(expectedLeft)} y:${Math.round(expectedTop)}`);
        console.log(`Actual: x:${Math.round(actualLeft)} y:${Math.round(actualTop)}`);
        console.log(`Difference: x:${leftDiff.toFixed(1)} y:${topDiff.toFixed(1)}`);
        
        if (topDiff > 10 || leftDiff > 10) {
            console.error(`SIGNIFICANT POSITION ERROR: Expected(${Math.round(expectedLeft)}, ${Math.round(expectedTop)}) vs Actual(${Math.round(actualLeft)}, ${Math.round(actualTop)})`);
            
            this.forceCorrectPosition(expectedTop, expectedLeft);
        }
    }
    
    private forceCorrectPosition(expectedTop: number, expectedLeft: number): void {
        if (!this._tooltipElement) return;
        
        console.log('Applying aggressive position correction...');
        
        this._tooltipElement.style.cssText = '';
        
        const correctionStyles = `
            position: fixed !important;
            top: ${Math.round(expectedTop)}px !important;
            left: ${Math.round(expectedLeft)}px !important;
            right: auto !important;
            bottom: auto !important;
            transform: none !important;
            translate: none !important;
            margin: 0 !important;
            z-index: 2147483647 !important;
            display: block !important;
            visibility: visible !important;
            opacity: 1 !important;
        `;
        
        this._tooltipElement.style.cssText = correctionStyles;
        
        void this._tooltipElement.offsetHeight;
        void this._tooltipElement.offsetWidth;
        
        setTimeout(() => {
            if (!this._tooltipElement) return;
            
            const finalRect = this._tooltipElement.getBoundingClientRect();
            const finalTop = Math.round(finalRect.top);
            const finalLeft = Math.round(finalRect.left);
            
            console.log(`After aggressive correction: x:${finalLeft} y:${finalTop}`);
            
            const remainingTopDiff = Math.abs(finalTop - expectedTop);
            const remainingLeftDiff = Math.abs(finalLeft - expectedLeft);
            
            if (remainingTopDiff > 5 || remainingLeftDiff > 5) {
                console.error('Position correction failed. This may indicate CSS conflicts or browser-specific issues.');
                console.log('Consider checking for:');
                console.log('1. CSS transforms on parent elements');
                console.log('2. CSS containment properties');
                console.log('3. Iframe scaling or zoom');
                console.log('4. Browser-specific positioning bugs');
                
                this.tryRelativeToBodyPositioning(expectedTop, expectedLeft);
            }
        }, 100);
    }
    
    private tryRelativeToBodyPositioning(expectedTop: number, expectedLeft: number): void {
        if (!this._tooltipElement) return;
        
        console.log('Trying body-relative positioning as last resort...');
        
        if (this._tooltipElement.parentNode && this._tooltipElement.parentNode !== document.body) {
            document.body.appendChild(this._tooltipElement);
        }
        
        const bodyRelativeStyles = `
            position: fixed !important;
            top: ${Math.round(expectedTop)}px !important;
            left: ${Math.round(expectedLeft)}px !important;
            right: auto !important;
            bottom: auto !important;
            transform: none !important;
            z-index: 2147483647 !important;
            margin: 0 !important;
            padding: 12px 16px !important;
            background-color: #333333 !important;
            color: #ffffff !important;
            border-radius: 4px !important;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3) !important;
            font-size: 13px !important;
            line-height: 1.5 !important;
            max-width: 400px !important;
            min-width: 200px !important;
            word-wrap: break-word !important;
            white-space: normal !important;
            visibility: visible !important;
            opacity: 1 !important;
            display: block !important;
        `;
        
        this._tooltipElement.style.cssText = bodyRelativeStyles;
        
        void this._tooltipElement.offsetHeight;
        
        setTimeout(() => {
            if (!this._tooltipElement) return;
            const rect = this._tooltipElement.getBoundingClientRect();
            console.log(`Body-relative positioning result: x:${Math.round(rect.left)} y:${Math.round(rect.top)}`);
        }, 50);
    }

    private getStableTooltipRect(): DOMRect | null {
        if (!this._tooltipElement) return null;
        
        const originalDisplay = this._tooltipElement.style.display;
        const originalVisibility = this._tooltipElement.style.visibility;
        const originalPosition = this._tooltipElement.style.position;
        
        this._tooltipElement.style.display = 'block';
        this._tooltipElement.style.visibility = 'hidden';
        this._tooltipElement.style.position = 'fixed';
        
        void this._tooltipElement.offsetHeight;
        void this._tooltipElement.offsetWidth;
        
        const rect = this._tooltipElement.getBoundingClientRect();
        
        this._tooltipElement.style.display = originalDisplay;
        this._tooltipElement.style.visibility = originalVisibility;
        this._tooltipElement.style.position = originalPosition;
        
        if (rect.width === 0 || rect.height === 0) {
            console.warn('Tooltip has zero dimensions');
            return null;
        }
        
        return rect;
    }

    private applyFallbackPosition(): void {
        if (!this._tooltipElement) return;
        
        console.warn('Applying fallback positioning');
        
        this._tooltipElement.style.setProperty('position', 'fixed', 'important');
        this._tooltipElement.style.setProperty('top', '50px', 'important');
        this._tooltipElement.style.setProperty('left', '50px', 'important');
        this._tooltipElement.style.setProperty('zIndex', '2147483647', 'important');
        this._tooltipElement.style.setProperty('maxWidth', '300px', 'important');
    }

    private handleResize(): void {
        if (this._tooltipElement && this._tooltipElement.classList.contains('visible')) {
            this.positionTooltip();
        }
    }

    private addWindowResizeHandler(): void {
        window.addEventListener('resize', this._resizeHandlerBound);
        window.addEventListener('scroll', this._resizeHandlerBound);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;

        const isVisible = true;
        if (this._container) {
            this._container.style.display = isVisible ? 'inline-block' : 'none';
        }

        this.updateRedirectFunctionality();
        this.updateAccessibilityAttributes();

        this.updateStyles();

        this.updateTooltipContent();

        if (this._iconElement) {
            const iconColor = this.getStringParameter('iconColor', '#666666');
            const iconSize = this.getNumberParameter('iconSize', 16);
            
            this._iconElement.style.color = iconColor;
            this._iconElement.style.fontSize = `${iconSize}px`;
            
            this.updateIconContent();
        }

        if (this._tooltipElement && this._tooltipElement.classList.contains('visible')) {
            this.positionTooltip();
        }
    }

    private updateStyles(): void {
        if (!this._styleElement) return;

        const bgColor = this.getStringParameter('tooltipBackgroundColor', '#333333');
        const textColor = this.getStringParameter('tooltipTextColor', '#ffffff');
        
        if (bgColor === this._lastBgColor && textColor === this._lastTextColor) {
            return;
        }
        
        this._lastBgColor = bgColor;
        this._lastTextColor = textColor;
        
        const linkColor = this.getLinkColor(textColor);
        
        document.documentElement.style.setProperty('--tooltip-bg-color', bgColor);
        document.documentElement.style.setProperty('--tooltip-text-color', textColor);
        document.documentElement.style.setProperty('--tooltip-link-color', linkColor);
        
        this._styleElement.textContent = this.generateEnhancedStyleContent(bgColor, textColor, linkColor);

        if (this._tooltipElement) {
            this._tooltipElement.style.setProperty('background-color', bgColor, 'important');
            this._tooltipElement.style.setProperty('background', bgColor, 'important');
            this._tooltipElement.style.setProperty('color', textColor, 'important');
            this._tooltipElement.style.setProperty('filter', 'none', 'important');
            this._tooltipElement.style.setProperty('backdrop-filter', 'none', 'important');
        }
    }


    public getOutputs(): IOutputs {
        return {
            dummyField: this.getStringParameter('dummyField') || ""
        };
    }

    public destroy(): void {
        if (this._iconElement) {
            this._iconElement.removeEventListener('mouseenter', this._showTooltipBound);
            this._iconElement.removeEventListener('mouseleave', this._hideTooltipBound);
            this._iconElement.removeEventListener('mouseover', this._showTooltipBound);
            this._iconElement.removeEventListener('mouseout', this._hideTooltipBound);
            this._iconElement.removeEventListener('focus', this._showTooltipBound);
            this._iconElement.removeEventListener('blur', this._hideTooltipBound);
            this._iconElement.removeEventListener('click', this._clickHandlerBound);
        }

        window.removeEventListener('resize', this._resizeHandlerBound);
        window.removeEventListener('scroll', this._resizeHandlerBound);
        
        if (this._hideTimeout) {
            clearTimeout(this._hideTimeout);
            this._hideTimeout = null;
        }
        
        if (this._positionUpdateFrameId) {
            cancelAnimationFrame(this._positionUpdateFrameId);
            this._positionUpdateFrameId = null;
        }
        
        if (this._styleElement && this._styleElement.parentNode) {
            this._styleElement.parentNode.removeChild(this._styleElement);
        }

        if (this._tooltipElement && this._tooltipElement.parentNode === document.body) {
            document.body.removeChild(this._tooltipElement);
        }

        this._container = null!;
        this._context = null!;
        this._tooltipElement = null!;
        this._iconElement = null!;
        this._styleElement = null!;
        this._notifyOutputChanged = null!;
    }
}