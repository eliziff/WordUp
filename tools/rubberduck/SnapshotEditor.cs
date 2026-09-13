// SPDX-License-Identifier: GPL-3.0-or-later
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace WordUp.Analysis {
    // Rubberduck's editor contracts include interactive and COM mutation methods.
    // This read-only boundary supports only explicitly supplied snapshot operations;
    // everything else fails, never returns a plausible default or touches Word.
    internal sealed class SnapshotBoundary : RealProxy {
        readonly Dictionary<string, Func<object[], object>> operations;
        SnapshotBoundary(Type type, Dictionary<string, Func<object[], object>> operations) : base(type) {
            this.operations = operations;
        }
        public static T Create<T>(Dictionary<string, Func<object[], object>> operations) where T : class {
            return (T)new SnapshotBoundary(typeof(T), operations).GetTransparentProxy();
        }
        public override IMessage Invoke(IMessage message) {
            var call = (IMethodCallMessage)message;
            try {
                if (!operations.TryGetValue(call.MethodName, out var operation))
                    throw new NotSupportedException("Workspace snapshot does not support " + call.MethodBase.DeclaringType.Name + "." + call.MethodName);
                return new ReturnMessage(operation(call.Args), null, 0, call.LogicalCallContext, call);
            } catch (Exception error) { return new ReturnMessage(error, call); }
        }
    }

    internal sealed class SnapshotEditor : IProjectsRepository {
        readonly IVBProject project;
        readonly string projectId;
        readonly Dictionary<QualifiedModuleName, IVBComponent> components = new Dictionary<QualifiedModuleName, IVBComponent>();
        public readonly Dictionary<QualifiedModuleName, SourceSnapshot> Sources = new Dictionary<QualifiedModuleName, SourceSnapshot>();
        public IVBE Vbe { get; }
        public IVbeEvents Events { get; }
        public SnapshotEditor(string projectId, string projectName, IEnumerable<(string Name, ComponentType Type, SourceSnapshot Source)> modules) {
            if (String.IsNullOrWhiteSpace(projectId) || String.IsNullOrWhiteSpace(projectName)) throw new ArgumentException("Project identity required");
            this.projectId = projectId;
            Vbe = SnapshotBoundary.Create<IVBE>(new Dictionary<string, Func<object[], object>> {
                {"HostApplication", _ => null}, {"get_IsInDesignMode", _ => true}, {"get_Version", _ => "7.1"}
            });
            var events = new Dictionary<string, Func<object[], object>>();
            // Immutable snapshots cannot emit live editor events; subscription lifetime is a no-op.
            foreach (var name in new[]{"ProjectAdded","ProjectRemoved","ProjectRenamed","ComponentAdded","ComponentRemoved","ComponentRenamed"}) {
                events.Add("add_" + name, _ => null); events.Add("remove_" + name, _ => null);
            }
            Events = SnapshotBoundary.Create<IVbeEvents>(events);
            project = SnapshotBoundary.Create<IVBProject>(new Dictionary<string, Func<object[], object>> {
                {"get_Name", _ => projectName}, {"get_ProjectId", _ => projectId},
                {"get_IsWrappingNullReference", _ => false}, {"get_Protection", _ => ProjectProtection.Unprotected},
                {"Dispose", _ => null}
            });
            var collection = SnapshotBoundary.Create<IVBComponents>(new Dictionary<string, Func<object[], object>> {
                {"get_Parent", _ => project}, {"Dispose", _ => null}
            });
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var module in modules) {
                if (!names.Add(module.Name)) throw new ArgumentException("Duplicate VBA module: " + module.Name);
                if (module.Type != ComponentType.StandardModule && module.Type != ComponentType.ClassModule)
                    throw new NotSupportedException("Designer/document metadata adapter required: " + module.Name);
                var source = module.Source ?? throw new ArgumentException("Module source required");
                var component = SnapshotBoundary.Create<IVBComponent>(new Dictionary<string, Func<object[], object>> {
                    {"get_Name", _ => module.Name}, {"get_Type", _ => module.Type},
                    {"get_IsWrappingNullReference", _ => false}, {"get_Collection", _ => collection},
                    {"get_HasDesigner", _ => false}, {"ContentHash", _ => source.CodePane.GetHashCode()},
                    {"Dispose", _ => null}
                });
                var key = new QualifiedModuleName(component);
                components.Add(key, component); Sources.Add(key, source);
            }
        }
        public IVBProject Project(string id) {
            if (id != projectId) throw new ArgumentException("Unknown snapshot project: " + id);
            return project;
        }
        public IEnumerable<(string ProjectId, IVBProject Project)> Projects() { yield return (projectId, project); }
        public IEnumerable<(string ProjectId, IVBProject Project)> LockedProjects() { yield break; }
        public IEnumerable<(QualifiedModuleName QualifiedModuleName, IVBComponent Component)> Components() => components.Select(p => (p.Key,p.Value));
        public IEnumerable<(QualifiedModuleName QualifiedModuleName, IVBComponent Component)> Components(string id) { Project(id); return Components(); }
        public IVBComponent Component(QualifiedModuleName key) => components.TryGetValue(key, out var component) ? component : null;
        public void Refresh() { }
        public void Refresh(string id) { Project(id); }
        public void Dispose() { }
        public IVBProjects ProjectsCollection() => throw new NotSupportedException("No live project collection");
        public IVBComponents ComponentsCollection(string id) => throw new NotSupportedException("No live component collection");
        public void RemoveComponent(QualifiedModuleName key) => throw new NotSupportedException("Snapshot is immutable");
    }
}
