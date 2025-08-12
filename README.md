# Power Platform Tooltip Control

A sophisticated and feature-rich tooltip control for Microsoft Power Platform applications, providing enhanced user experience through contextual help and information display.

## Features

### Core Functionality
- **Hover Tooltips**: Display rich tooltips on icon hover/focus
- **Click Navigation**: Optional redirect functionality with URL navigation
- **Smart Positioning**: Intelligent tooltip positioning that adapts to viewport boundaries
- **Cross-Platform**: Works seamlessly across Power Platform environments (Canvas Apps, Model-Driven Apps, Power Pages)

### Customization Options

#### Visual Appearance
- **Custom Icons**: Choose from built-in icon types (info, question, warning, error) or use custom symbols
- **Flexible Styling**: 
  - Configurable icon color and size
  - Customizable tooltip background and text colors
  - Automatic link color generation based on theme
  - Responsive design with mobile optimization

#### Content Management
- **Rich Content Support**: 
  - HTML content rendering with sanitization
  - Plain text with line break preservation
  - Support for headings, lists, images, and links
  - Formatted text (bold, italic, underline)
- **Dynamic Content**: Bind tooltip content to data sources or use static text
- **Content Security**: Built-in HTML sanitization to prevent XSS attacks

#### Positioning & Layout
- **Smart Positioning**: Automatic positioning (top, bottom, left, right) based on available space
- **Manual Positioning**: Override automatic positioning with preferred placement
- **Viewport Awareness**: Automatically adjusts to stay within screen boundaries
- **Power Platform Optimized**: Special handling for iframe contexts and form layouts
- **Responsive Sizing**: Configurable maximum width and height with mobile breakpoints

#### Navigation Features
- **URL Redirection**: Click-to-navigate functionality
- **Target Control**: Choose to open links in new tab or current window
- **URL Validation**: Built-in URL sanitization and validation
- **Protocol Support**: Supports HTTP/HTTPS, relative URLs, mailto, and tel links

#### Accessibility
- **ARIA Compliance**: Full ARIA attributes for screen readers
- **Keyboard Navigation**: Tab and focus support
- **High Contrast**: Respects system accessibility settings
- **Screen Reader Support**: Descriptive labels and roles

##  Configuration Properties

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `staticTooltipContent` | String | "Tooltip content" | Static text content for the tooltip |
| `boundTooltipContent` | String | - | Data-bound content source |
| `iconType` | String | "info" | Icon type: info, question, warning, error |
| `iconColor` | String | "#666666" | Color of the tooltip icon |
| `iconSize` | Number | 16 | Size of the icon in pixels |
| `tooltipBackgroundColor` | String | "#333333" | Background color of the tooltip |
| `tooltipTextColor` | String | "#ffffff" | Text color inside the tooltip |
| `position` | String | "right" | Preferred tooltip position: top, bottom, left, right |
| `allowHtml` | Boolean | false | Enable HTML content rendering |
| `preserveLineBreaks` | Boolean | true | Convert line breaks to `<br>` tags |
| `redirectUrl` | String | - | URL to navigate to on icon click |
| `openInNewTab` | Boolean | true | Open redirect URL in new tab |
| `tooltipMaxWidth` | Number | 400 | Maximum width of tooltip in pixels |
| `tooltipMaxHeight` | Number | - | Maximum height of tooltip in pixels |

## Use Cases

### Help & Documentation
- **Field Explanations**: Provide detailed explanations for complex form fields
- **Process Guidance**: Step-by-step instructions for business processes
- **Data Definitions**: Explain business terms and data meanings
- **Validation Rules**: Inform users about input requirements and constraints

### Navigation & Workflows
- **Quick Links**: Provide shortcuts to related records or external resources
- **Process Navigation**: Guide users through multi-step workflows
- **Reference Materials**: Link to documentation, policies, or training materials
- **External Integrations**: Connect to external systems or applications

### Rich Information Display
- **Contextual Help**: Show relevant information without cluttering the interface
- **Preview Information**: Display summary data before navigation
- **Status Explanations**: Explain the meaning of status codes or indicators
- **Formula Explanations**: Describe calculation logic in user-friendly terms

## Technical Features

### Performance Optimizations
- **GPU Acceleration**: Smooth animations and positioning
- **Memory Management**: Proper cleanup and event handling
- **Frame Rate Optimization**: RequestAnimationFrame for smooth interactions
- **Layout Stability**: Prevents layout shifts and content jumping

### Power Platform Integration
- **Iframe Support**: Handles Power Platform's iframe architecture
- **Form Integration**: Optimized for model-driven app forms
- **Canvas App Compatible**: Works seamlessly in Canvas Apps
- **Power Pages Ready**: Functions correctly in Power Pages sites

### Browser Compatibility
- **Modern Browsers**: Chrome, Firefox, Safari, Edge
- **Mobile Responsive**: Touch-friendly interactions
- **Cross-Device**: Consistent experience across devices
- **Accessibility Standards**: WCAG 2.1 compliant

### Security Features
- **XSS Prevention**: Comprehensive HTML sanitization
- **URL Validation**: Safe handling of redirect URLs
- **Content Filtering**: Removes dangerous scripts and content
- **CORS Handling**: Proper handling of cross-origin scenarios

## Responsive Design

### Mobile Optimization
- **Touch Interactions**: Optimized for touch devices
- **Viewport Scaling**: Automatic sizing based on screen size
- **Gesture Support**: Tap to show/hide on mobile devices
- **Performance**: Lightweight and fast on mobile browsers

### Breakpoint Handling
- **Small Screens** (< 320px): Minimal padding and sizing
- **Mobile** (< 480px): Compact layout
- **Tablet & Desktop**: Full feature set with optimal spacing

## Installation & Setup

1. **Import the Solution**: Import the managed solution into your Power Platform environment
2. **Add to Forms**: Drag the control onto your forms or canvas apps
3. **Configure Properties**: Set up the tooltip content and styling options
4. **Test Functionality**: Verify tooltip display and navigation features

## Styling Examples

### Basic Information Tooltip
```
Icon Type: info
Icon Color: #0078d4
Tooltip Background: #f3f2f1
Tooltip Text: #323130
Content: "This field is required for processing your request."
```

### Warning with Navigation
```
Icon Type: warning
Icon Color: #ff8c00
Tooltip Background: #fff4ce
Tooltip Text: #323130
Content: "Data validation failed. Click for troubleshooting guide."
Redirect URL: https://docs.company.com/troubleshooting
```

### Rich HTML Content
```
Allow HTML: true
Content: "<h3>Account Status</h3><ul><li><strong>Active:</strong> Account is operational</li><li><strong>Pending:</strong> Awaiting approval</li><li><strong>Suspended:</strong> Temporarily disabled</li></ul>"
```

## Security Considerations

- All HTML content is sanitized to prevent XSS attacks
- URL validation prevents malicious redirects
- Content Security Policy (CSP) compatible
- No external dependencies reduce attack surface

## Performance Metrics

- **Load Time**: < 50ms initialization
- **Memory Usage**: < 2MB memory footprint
- **Render Time**: < 16ms for smooth 60fps animations
- **Bundle Size**: Optimized for minimal impact

## Contributing

This control is designed for enterprise use in Power Platform environments. For customizations or enhancements, please follow your organization's development guidelines.

## License

This control is provided as-is for use within Microsoft Power Platform environments. Please ensure compliance with your organization's licensing requirements.

---

**Built for Microsoft Power Platform** | **Enterprise Ready** | **Accessibility Compliant**
