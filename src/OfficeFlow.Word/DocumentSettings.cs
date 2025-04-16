using System.IO;

namespace OfficeFlow.Word;

public sealed class DocumentSettings
{
    /// <summary>
    /// Default settings
    /// </summary>
    public static DocumentSettings Default
        => new()
        {
            IsReadOnly = false,
            AllowAutoSaving = false,
            AllowUpdateFieldsOnOpen = true
        };

    /// <inheritdoc cref="AllowAutoSaving"/>
    private bool _allowAutoSaving;


    /// <summary>
    /// Document access mode
    /// </summary>
    internal FileAccess AccessMode
        => IsReadOnly switch
        {
            true => FileAccess.Read,
            false => FileAccess.ReadWrite
        };

    /// <summary>
    /// Open a document in read-only mode
    /// </summary>
    public bool IsReadOnly { get; set; }

    /// <summary>
    /// Allow automatic saving when closing a document
    /// </summary>
    /// <remarks>
    /// In read-only mode, this setting will be automatically disabled
    /// </remarks>
    public bool AllowAutoSaving
    {
        get => _allowAutoSaving && !IsReadOnly;
        set => _allowAutoSaving = value;
    }

    /// <summary>
    /// Updating fields when opening a document
    /// </summary>
    public bool AllowUpdateFieldsOnOpen { get; set; }
}