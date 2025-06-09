using DrillNamer.Core;
using NUnit.Framework;

namespace DrillNamer.Tests;

public class LayerStateTests
{
    private class FakeLayer : ILayerLock
    {
        public bool IsLocked { get; set; } = true;
        public string Name => "FAKE";
    }

    [Test]
    public void WithUnlocked_RelocksLayer()
    {
        var layer = new FakeLayer { IsLocked = true };
        bool ran = false;
        LayerState.WithUnlocked(layer, () => ran = true);
        Assert.IsTrue(ran);
        Assert.IsTrue(layer.IsLocked);
    }
}
