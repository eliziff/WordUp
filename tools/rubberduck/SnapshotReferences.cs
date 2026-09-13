// SPDX-License-Identifier: GPL-3.0-or-later
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.ComReferenceLoading;
using Rubberduck.VBEditor;

namespace WordUp.Analysis {
    internal sealed class SnapshotReferences : IProjectReferencesProvider {
        // Rubberduck caches reflected types statically. Serialize reflection and
        // clear between batches so another library version cannot poison a run.
        static readonly object ReflectionLock = new object();
        readonly List<ComProject> projects = new List<ComProject>();
        public int DeclarationCount { get; }
        // Resolver declarations are mutable (references and type bindings). Never
        // reuse them across analyses; only the reflected library model is retained.
        public IEnumerable<Declaration> CreateDeclarations() => projects.SelectMany(p => new DeclarationsFromComProjectLoader().LoadDeclarations(p));
        public IReadOnlyCollection<ReferencePriorityMap> ProjectReferences { get; }
        public IReadOnlyCollection<(ReferenceInfo Reference, string Error)> Unavailable { get; }
        public IReadOnlyDictionary<string,string> Hashes { get; }
        public SnapshotReferences(string projectId, IReadOnlyList<ReferenceInfo> references) {
            if (references == null || references.Count > 128) throw new ArgumentException("Reference budget exceeded");
            var declarations = new List<Declaration>();
            var priorities = new List<ReferencePriorityMap>();
            var unavailable = new List<(ReferenceInfo,string)>();
            var hashes = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            lock (ReflectionLock) {
                ComProject.KnownTypes.Clear(); ComProject.KnownAliases.Clear(); ComProject.KnownEnumerations.Clear();
                for (int i=0; i<references.Count; i++) {
                    var reference = references[i];
                    try {
                        if (!Path.IsPathRooted(reference.FullPath)) throw new ArgumentException("Absolute type-library path required");
                        // Keep the file read-locked across reflection and hashing.
                        using (var stream = new FileStream(reference.FullPath,FileMode.Open,FileAccess.Read,FileShare.Read)) {
                            if (stream.Length > 256L*1024*1024) throw new IOException("Type-library byte budget exceeded");
                            var provider = new ComLibraryProvider();
                            var library = provider.LoadTypeLibrary(reference.FullPath);
                            if (library == null) throw new InvalidOperationException("Type library could not be loaded");
                            try {
                                var actual = provider.GetReferenceInfo(library,reference.Name,reference.FullPath);
                                if (actual.Guid != reference.Guid || actual.Major != reference.Major || actual.Minor != reference.Minor)
                                    throw new InvalidOperationException("Type-library GUID/version does not match the workspace reference");
                                var reflected = new ComProject(library,reference.FullPath);
                                if (!String.Equals(reflected.Name,reference.Name,StringComparison.OrdinalIgnoreCase))
                                    throw new InvalidOperationException("Type-library name does not match the workspace reference");
                                var loaded = new DeclarationsFromComProjectLoader().LoadDeclarations(reflected);
                                var project = loaded.OfType<ProjectDeclaration>().Single();
                                if (priorities.Any(p => p.ReferencedProjectId == project.ProjectId))
                                    throw new ArgumentException("Duplicate type-library reference");
                                string digest;
                                using (var hash = SHA256.Create())
                                    digest = BitConverter.ToString(hash.ComputeHash(stream)).Replace("-","").ToLowerInvariant();
                                hashes.Add(reference.FullPath,digest);
                                declarations.AddRange(loaded);
                                projects.Add(reflected);
                                priorities.Add(new ReferencePriorityMap(project.ProjectId){{projectId,i+1}});
                            } finally { Marshal.ReleaseComObject(library); }
                        }
                    } catch (Exception error) when (error is IOException || error is UnauthorizedAccessException ||
                        error is COMException || error is ArgumentException || error is InvalidOperationException) {
                        unavailable.Add((reference,error.Message));
                    }
                }
            }
            DeclarationCount = declarations.Count; ProjectReferences = priorities; Unavailable = unavailable; Hashes = hashes;
        }
    }
}
